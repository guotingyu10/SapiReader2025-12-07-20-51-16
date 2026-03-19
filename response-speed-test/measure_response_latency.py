import argparse
import ctypes
import ctypes.wintypes as wintypes
import json
import os
import random
import socket
import subprocess
import sys
import threading
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Optional, Tuple


VOICE_BASELINES_MS: Dict[str, float] = {
    "Microsoft Jenny": 62.5,
    "Microsoft Xiaoxiao": 62.9,
}


def _try_import_audio_deps():
    try:
        import numpy as np  # type: ignore
        import sounddevice as sd  # type: ignore
    except Exception:
        print(
            "缺少依赖：sounddevice / numpy\n"
            "请先安装：\n"
            "  python -m pip install sounddevice numpy\n"
            "然后重试。",
            file=sys.stderr,
        )
        raise
    return sd, np


def find_config_path(start_dir: Path) -> Path:
    for p in [start_dir, *start_dir.parents]:
        candidate = p / "config.json"
        if candidate.exists():
            return candidate
    raise FileNotFoundError("未找到 config.json（从脚本目录向上查找）")


def load_json(path: Path) -> Dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def save_json(path: Path, data: Dict[str, Any]) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def _rand_cn_chunk(min_len: int, max_len: int) -> str:
    n = random.randint(min_len, max_len)
    return "".join(chr(random.randint(0x4E00, 0x9FA5)) for _ in range(n))


def _rand_ascii_token(min_len: int, max_len: int) -> str:
    n = random.randint(min_len, max_len)
    letters = "bcdfghjklmnpqrstvwxyz"
    vowels = "aeiou"
    out = []
    for i in range(n):
        if i % 2 == 0:
            out.append(random.choice(letters))
        else:
            out.append(random.choice(vowels))
    if random.random() < 0.35:
        out.insert(random.randint(0, len(out)), str(random.randint(0, 9999)))
    token = "".join(out)
    if random.random() < 0.25:
        token = token.upper()
    return token


def build_random_mixed_text(min_chars: int) -> str:
    parts = []
    total = 0
    while total < min_chars:
        if random.random() < 0.55:
            s = _rand_cn_chunk(3, 10)
        else:
            s = _rand_ascii_token(4, 14)
        parts.append(s)
        total += len(s)
        sep = random.choice([" ", " ", " ", ",", ";", ":", ".", " / "])
        parts.append(sep)
        total += len(sep)
    text = "".join(parts).strip()
    if len(text) > min_chars:
        text = text[:min_chars]
    return text


def build_switch_text(start_with_chinese: bool, total_chars: int = 120) -> str:
    cn = _rand_cn_chunk(8, 16)
    en = _rand_ascii_token(10, 24)
    mid = build_random_mixed_text(max(40, total_chars - len(cn) - len(en) - 10))
    if start_with_chinese:
        text = f"{cn} {en} {mid}"
    else:
        text = f"{en} {cn} {mid}"
    if len(text) > total_chars:
        text = text[:total_chars]
    return text.strip()


def create_short_text_for_voice(voice_name: str) -> str:
    if "Xiaoxiao" in voice_name:
        return "你好，这是一次延迟测试。"
    return "Hello, this is a latency test."


GMEM_MOVEABLE = 0x0002
GMEM_ZEROINIT = 0x0040
CF_UNICODETEXT = 13


def set_clipboard_text(text: str) -> None:
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32

    user32.OpenClipboard.restype = wintypes.BOOL
    user32.OpenClipboard.argtypes = [wintypes.HWND]
    user32.CloseClipboard.restype = wintypes.BOOL
    user32.CloseClipboard.argtypes = []
    user32.EmptyClipboard.restype = wintypes.BOOL
    user32.EmptyClipboard.argtypes = []
    user32.SetClipboardData.restype = wintypes.HANDLE
    user32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]

    kernel32.GlobalAlloc.restype = wintypes.HGLOBAL
    kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
    kernel32.GlobalFree.restype = wintypes.HGLOBAL
    kernel32.GlobalFree.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalLock.restype = ctypes.c_void_p
    kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalUnlock.restype = wintypes.BOOL
    kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]

    for _ in range(50):
        if user32.OpenClipboard(None):
            break
        time.sleep(0.02)
    else:
        raise RuntimeError("OpenClipboard failed")

    try:
        if not user32.EmptyClipboard():
            raise RuntimeError("EmptyClipboard failed")
        data = (text + "\0").encode("utf-16le")
        h_global = kernel32.GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, len(data))
        if not h_global:
            raise RuntimeError("GlobalAlloc failed")
        locked = kernel32.GlobalLock(h_global)
        if not locked:
            kernel32.GlobalFree(h_global)
            raise RuntimeError("GlobalLock failed")
        try:
            ctypes.memmove(locked, data, len(data))
        finally:
            kernel32.GlobalUnlock(h_global)
        if not user32.SetClipboardData(CF_UNICODETEXT, h_global):
            kernel32.GlobalFree(h_global)
            raise RuntimeError("SetClipboardData failed")
    finally:
        user32.CloseClipboard()


INPUT_KEYBOARD = 1
KEYEVENTF_KEYUP = 0x0002


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", wintypes.WORD),
        ("wScan", wintypes.WORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
    ]


class _INPUT_UNION(ctypes.Union):
    _fields_ = [("ki", KEYBDINPUT)]


class INPUT(ctypes.Structure):
    _fields_ = [("type", wintypes.DWORD), ("u", _INPUT_UNION)]


def _send_vk(vk: int, key_up: bool) -> None:
    extra = ctypes.c_ulong(0)
    flags = KEYEVENTF_KEYUP if key_up else 0
    inp = INPUT(type=INPUT_KEYBOARD, u=_INPUT_UNION(ki=KEYBDINPUT(vk, 0, flags, 0, ctypes.pointer(extra))))
    ctypes.windll.user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))


def send_key(vk: int) -> None:
    _send_vk(vk, False)
    _send_vk(vk, True)


def send_chord(vk_mod: int, vk_key: int) -> None:
    _send_vk(vk_mod, False)
    _send_vk(vk_key, False)
    _send_vk(vk_key, True)
    _send_vk(vk_mod, True)


VK_CONTROL = 0x11
VK_A = 0x41
VK_V = 0x56
VK_F1 = 0x70
VK_F3 = 0x72
VK_SHIFT = 0x10
VK_HOME = 0x24
VK_END = 0x23


EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)


def find_main_window_by_pid(pid: int) -> Optional[int]:
    user32 = ctypes.windll.user32
    matches = []

    def callback(hwnd: int, lparam: int) -> bool:
        if not user32.IsWindowVisible(hwnd):
            return True
        proc_id = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(proc_id))
        if proc_id.value == pid:
            matches.append(hwnd)
            return False
        return True

    user32.EnumWindows(EnumWindowsProc(callback), 0)
    return matches[0] if matches else None


def activate_window(hwnd: int) -> None:
    user32 = ctypes.windll.user32
    user32.ShowWindow(hwnd, 5)
    user32.SetForegroundWindow(hwnd)


def key_down(vk: int) -> None:
    _send_vk(vk, False)


def key_up(vk: int) -> None:
    _send_vk(vk, True)


def select_all_in_window(hwnd: int) -> None:
    activate_window(hwnd)
    time.sleep(0.08)
    send_chord(VK_CONTROL, VK_A)
    time.sleep(0.05)


def select_first_line_in_window(hwnd: int) -> None:
    activate_window(hwnd)
    time.sleep(0.08)
    key_down(VK_CONTROL)
    send_key(VK_HOME)
    key_up(VK_CONTROL)
    time.sleep(0.03)
    key_down(VK_SHIFT)
    send_key(VK_END)
    key_up(VK_SHIFT)
    time.sleep(0.05)


def get_window_text(hwnd: int) -> str:
    user32 = ctypes.windll.user32
    user32.GetWindowTextLengthW.restype = ctypes.c_int
    user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
    user32.GetWindowTextW.restype = ctypes.c_int
    user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]

    n = user32.GetWindowTextLengthW(hwnd)
    if n <= 0:
        return ""
    buf = ctypes.create_unicode_buffer(n + 1)
    user32.GetWindowTextW(hwnd, buf, n + 1)
    return buf.value


def find_top_window_by_title_substring(title_substring: str) -> Optional[int]:
    user32 = ctypes.windll.user32
    matches = []
    needle = title_substring.lower()

    def callback(hwnd: int, lparam: int) -> bool:
        if not user32.IsWindowVisible(hwnd):
            return True
        title = get_window_text(hwnd)
        if title and needle in title.lower():
            matches.append(hwnd)
            return False
        return True

    user32.EnumWindows(EnumWindowsProc(callback), 0)
    return int(matches[0]) if matches else None


def open_text_window(file_path: Path, window_title_hint: str, editor_mode: str) -> Tuple[Optional[subprocess.Popen], int]:
    proc: Optional[subprocess.Popen] = None
    if editor_mode == "notepad":
        proc = subprocess.Popen(
            ["notepad.exe", str(file_path)],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        time.sleep(0.2)
    else:
        os.startfile(str(file_path))
        time.sleep(0.2)

    hwnd: Optional[int] = None
    if proc is not None:
        for _ in range(80):
            time.sleep(0.1)
            hwnd = find_main_window_by_pid(proc.pid)
            if hwnd:
                break

    if not hwnd:
        for _ in range(120):
            time.sleep(0.1)
            hwnd = find_top_window_by_title_substring(window_title_hint)
            if hwnd:
                break

    if not hwnd:
        if proc is not None:
            try:
                proc.terminate()
            except Exception:
                pass
        raise RuntimeError(f"无法找到文本窗口（标题不包含：{window_title_hint}）")

    activate_window(hwnd)
    time.sleep(0.25)
    return proc, hwnd


WM_SETTEXT = 0x000C
EM_SETSEL = 0x00B1


EnumChildWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)


def get_window_class_name(hwnd: int) -> str:
    buf = ctypes.create_unicode_buffer(256)
    ctypes.windll.user32.GetClassNameW(hwnd, buf, len(buf))
    return buf.value


def find_notepad_editor(hwnd: int) -> Optional[int]:
    user32 = ctypes.windll.user32
    user32.FindWindowExW.restype = wintypes.HWND
    user32.FindWindowExW.argtypes = [wintypes.HWND, wintypes.HWND, wintypes.LPCWSTR, wintypes.LPCWSTR]

    for cls in ["Edit", "RichEditD2DPT", "RICHEDIT50W", "RichEdit50W", "RichEdit20W"]:
        child = user32.FindWindowExW(hwnd, 0, cls, None)
        if child:
            return int(child)

    found: Optional[int] = None

    def callback(child_hwnd: int, lparam: int) -> bool:
        nonlocal found
        cls_name = get_window_class_name(child_hwnd)
        if "edit" in cls_name.lower():
            found = int(child_hwnd)
            return False
        return True

    user32.EnumChildWindows(hwnd, EnumChildWindowsProc(callback), 0)
    return found


def notepad_set_and_select_all(text: str, hwnd: int) -> None:
    activate_window(hwnd)
    time.sleep(0.08)
    editor = find_notepad_editor(hwnd)
    if editor:
        user32 = ctypes.windll.user32
        user32.SetFocus(editor)
        user32.SendMessageW(editor, WM_SETTEXT, 0, ctypes.c_wchar_p(text))
        user32.SendMessageW(editor, EM_SETSEL, 0, wintypes.LPARAM(-1))
        time.sleep(0.05)
        send_chord(VK_CONTROL, VK_A)
        time.sleep(0.03)
        return

    set_clipboard_text(text)
    time.sleep(0.05)
    send_chord(VK_CONTROL, VK_A)
    time.sleep(0.03)
    send_chord(VK_CONTROL, VK_V)
    time.sleep(0.06)
    send_chord(VK_CONTROL, VK_A)
    time.sleep(0.03)


@dataclass
class AudioDetector:
    sd: Any
    np: Any
    sample_rate: int
    threshold: float
    consecutive_blocks: int
    mode: str = "peak"

    armed: bool = False
    t0: float = 0.0
    detected_t: float = 0.0
    consecutive: int = 0
    done: threading.Event = field(default_factory=threading.Event)
    lock: threading.Lock = field(default_factory=threading.Lock)

    def arm(self, t0: float) -> None:
        with self.lock:
            self.armed = True
            self.t0 = t0
            self.detected_t = 0.0
            self.consecutive = 0
            self.done.clear()

    def disarm(self) -> None:
        with self.lock:
            self.armed = False
            self.consecutive = 0

    def set_threshold(self, threshold: float) -> None:
        with self.lock:
            self.threshold = float(threshold)

    def callback(self, indata, frames, time_info, status) -> None:
        with self.lock:
            if not self.armed or self.detected_t:
                return
            x = indata.astype(self.np.float32, copy=False)
            if self.mode == "rms":
                score = float(self.np.sqrt(self.np.mean(self.np.square(x))))
            else:
                score = float(self.np.max(self.np.abs(x)))

            if score >= self.threshold:
                self.consecutive += 1
                if self.consecutive >= self.consecutive_blocks:
                    self.detected_t = time.perf_counter()
                    self.done.set()
            else:
                if self.consecutive > 0:
                    self.consecutive -= 1

    def wait_detected(self, timeout_s: float) -> Optional[float]:
        ok = self.done.wait(timeout_s)
        if not ok:
            return None
        with self.lock:
            return self.detected_t


def calibrate_noise_threshold(
    sd: Any, np: Any, sample_rate: int, seconds: float, min_abs_threshold: float
) -> Tuple[float, float]:
    duration = max(0.2, seconds)
    frames = int(sample_rate * duration)
    audio = sd.rec(frames, samplerate=sample_rate, channels=1, dtype="float32", blocking=True)
    x = audio.astype(np.float32, copy=False)
    rms = float(np.sqrt(np.mean(np.square(x))))
    abs_x = np.abs(x).reshape(-1)
    p999 = float(np.quantile(abs_x, 0.999))
    threshold = max(float(min_abs_threshold), p999 + 0.02)
    if threshold > 0.08:
        threshold = 0.08
    return rms, threshold


def update_config_for_voice(config: Dict[str, Any], voice_name: str) -> Dict[str, Any]:
    updated = dict(config)
    updated["VoiceName"] = voice_name
    updated["AutoSwitchVoice"] = False
    updated["Volume"] = 100
    updated["HotkeyIndex"] = "F1"
    updated["StopHotkeyIndex"] = "F3"
    return updated


def api_connect(host: str, port: int) -> Tuple[socket.socket, Any]:
    last_err = None
    for i in range(10):
        try:
            s = socket.create_connection((host, port), timeout=3.0)
            s.setsockopt(socket.IPPROTO_TCP, socket.TCP_NODELAY, 1)
            f = s.makefile("rwb", buffering=0)
            return s, f
        except OSError as e:
            last_err = e
            time.sleep(0.4)
    raise RuntimeError(f"无法连接到控制 API {host}:{port}，请确认 SapiReader 已启动（{last_err})")


def api_call(f, obj: Dict[str, Any]) -> Dict[str, Any]:
    payload = (json.dumps(obj, ensure_ascii=False) + "\n").encode("utf-8")
    f.write(payload)
    line = f.readline()
    if not line:
        raise RuntimeError("API 连接已关闭")
    resp = json.loads(line.decode("utf-8", errors="replace"))
    if not isinstance(resp, dict):
        raise RuntimeError("API 返回非对象")
    if resp.get("ok") is False:
        raise RuntimeError(f"API 返回失败: {resp}")
    return resp


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser()
    p.add_argument("--api-host", default="127.0.0.1", help="SapiReader 控制 API host")
    p.add_argument("--api-port", type=int, default=32123, help="SapiReader 控制 API port")
    p.add_argument("--list-devices", action="store_true", help="列出音频输入设备并退出")
    p.add_argument("--audio-device", default=None, help="sounddevice 输入设备（index 或名称）")
    p.add_argument("--calibrate-seconds", type=float, default=0.6, help="每次测量前噪声校准时长")
    p.add_argument("--min-abs-threshold", type=float, default=0.002, help="最小绝对阈值（0~1）")
    p.add_argument("--consecutive-blocks", type=int, default=2, help="连续超阈值块数")
    p.add_argument("--blocksize", type=int, default=256, help="输入采样 blocksize（越小延迟越低）")
    p.add_argument("--input-latency", default="low", help="sounddevice InputStream latency（low/high/默认）")
    p.add_argument("--latency-compensate", action="store_true", help="用 InputStream.latency 进行延迟补偿")
    p.add_argument("--detect-mode", choices=["peak", "rms"], default="peak", help="出声检测方式")
    p.add_argument("--detect-timeout", type=float, default=5.0, help="等待音频启动的超时秒数")
    p.add_argument("--long-text-chars", type=int, default=2000, help="长文本字符数")
    p.add_argument("--pause-after-seconds", type=float, default=20.0, help="开始朗读后多少秒按停止")
    p.add_argument("--iterations", type=int, default=3, help="循环次数")
    return p.parse_args()


def main() -> int:
    args = parse_args()
    sd, np = _try_import_audio_deps()

    if args.list_devices:
        print(sd.query_devices())
        return 0

    if args.audio_device is not None:
        sd.default.device = args.audio_device

    try:
        sample_rate = int(sd.query_devices(device=args.audio_device, kind="input")["default_samplerate"])
    except Exception:
        sample_rate = 48000

    print(f"连接控制 API: {args.api_host}:{args.api_port}")
    sock, f = api_connect(args.api_host, int(args.api_port))
    try:
        results = []

        detector = AudioDetector(
            sd=sd,
            np=np,
            sample_rate=sample_rate,
            threshold=0.0,
            consecutive_blocks=max(1, int(args.consecutive_blocks)),
            mode=args.detect_mode,
        )

        stream = sd.InputStream(
            samplerate=sample_rate,
            channels=1,
            dtype="float32",
            callback=detector.callback,
            latency=None if args.input_latency in ("", "default", "none") else args.input_latency,
            blocksize=int(args.blocksize) if int(args.blocksize) > 0 else 0,
        )

        with stream:
            try:
                input_latency_s = float(stream.latency)
            except Exception:
                input_latency_s = 0.0
            api_call(f, {"cmd": "set_keepalive", "enable": False})
            api_call(f, {"cmd": "stop", "hardReset": False, "waitMs": 400})
            api_call(
                f,
                {
                    "cmd": "set_mixed_voices",
                    "chineseVoice": "Microsoft Xiaoxiao",
                    "englishVoice": "Microsoft Jenny",
                    "chineseRate": 10,
                    "englishRate": 5,
                    "volume": 100,
                    "mixedChineseMinChars": 1,
                    "mixedEnglishMinLetters": 1,
                    "warmUp": True,
                },
            )
            time.sleep(0.6)
            api_call(
                f,
                {
                    "cmd": "speak",
                    "text": build_switch_text(start_with_chinese=True, total_chars=80),
                    "skipProcessing": True,
                    "skipHistory": True,
                },
            )
            api_call(f, {"cmd": "stop", "hardReset": False, "waitMs": 600})

            for i in range(int(args.iterations)):
                print(f"循环 {i + 1}/{int(args.iterations)}")
                desired_len = max(int(args.long_text_chars), int(float(args.pause_after_seconds) * 400))
                long_text = build_random_mixed_text(desired_len)
                print("步骤1/2: 朗读大段中英文混合内容 -> 等待 -> 暂停")
                api_call(f, {"cmd": "speak", "text": long_text, "skipProcessing": True, "skipHistory": True})
                time.sleep(float(args.pause_after_seconds))
                api_call(f, {"cmd": "stop", "hardReset": False, "waitMs": 800})
                time.sleep(0.2)
                for _ in range(20):
                    st = api_call(f, {"cmd": "get_state"})
                    if st.get("voiceState") != "Speaking":
                        break
                    time.sleep(0.05)
                time.sleep(0.25)

                if i % 2 == 0:
                    next_text = build_switch_text(start_with_chinese=True, total_chars=140)
                    expected_voice = "Microsoft Xiaoxiao"
                else:
                    next_text = build_switch_text(start_with_chinese=False, total_chars=140)
                    expected_voice = "Microsoft Jenny"

                print("步骤2/2: 朗读下一个内容并测量出声延迟")
                noise_rms, threshold = calibrate_noise_threshold(
                    sd=sd,
                    np=np,
                    sample_rate=sample_rate,
                    seconds=float(args.calibrate_seconds),
                    min_abs_threshold=float(args.min_abs_threshold),
                )
                detector.set_threshold(threshold)

                t0 = time.perf_counter()
                detector.arm(t0)
                api_call(f, {"cmd": "speak", "text": next_text, "skipProcessing": True, "skipHistory": True})
                detected_t = detector.wait_detected(float(args.detect_timeout))
                detector.disarm()

                baseline_ms = VOICE_BASELINES_MS[expected_voice]
                limit_ms = baseline_ms * 3.0
                if detected_t is None:
                    results.append(
                        {
                            "expected_voice": expected_voice,
                            "iteration": i + 1,
                            "threshold": threshold,
                            "noise_rms": noise_rms,
                            "latency_ms": None,
                            "passed": False,
                            "limit_ms": limit_ms,
                            "baseline_ms": baseline_ms,
                            "raw_latency_ms": None,
                            "input_latency_ms": input_latency_s * 1000.0,
                        }
                    )
                    continue

                raw_latency_ms = (detected_t - t0) * 1000.0
                latency_ms = raw_latency_ms
                if args.latency_compensate and input_latency_s > 0:
                    latency_ms = max(0.0, raw_latency_ms - input_latency_s * 1000.0)
                passed = latency_ms <= limit_ms
                results.append(
                    {
                        "expected_voice": expected_voice,
                        "iteration": i + 1,
                        "threshold": threshold,
                        "noise_rms": noise_rms,
                        "latency_ms": latency_ms,
                        "raw_latency_ms": raw_latency_ms,
                        "input_latency_ms": input_latency_s * 1000.0,
                        "passed": passed,
                        "limit_ms": limit_ms,
                        "baseline_ms": baseline_ms,
                    }
                )

        all_passed = all(r["passed"] for r in results)
        print("结果：")
        for r in results:
            if r["latency_ms"] is None:
                print(
                    f'- Iter{r["iteration"]} {r["expected_voice"]}: 未检测到音频输出（noise_rms={r["noise_rms"]:.4f}, threshold={r["threshold"]:.4f}） -> FAIL (limit {r["limit_ms"]:.1f} ms)'
                )
            else:
                print(
                    f'- Iter{r["iteration"]} {r["expected_voice"]}: {r["latency_ms"]:.1f} ms (raw {r["raw_latency_ms"]:.1f} ms, input_latency {r["input_latency_ms"]:.1f} ms, baseline {r["baseline_ms"]:.1f} ms, limit {r["limit_ms"]:.1f} ms, noise_rms={r["noise_rms"]:.4f}, threshold={r["threshold"]:.4f}) -> '
                    + ("PASS" if r["passed"] else "FAIL")
                )
        return 0 if all_passed else 1
    finally:
        try:
            try:
                api_call(f, {"cmd": "set_keepalive", "enable": True})
            except Exception:
                pass
            f.close()
        except Exception:
            pass
        try:
            sock.close()
        except Exception:
            pass


if __name__ == "__main__":
    raise SystemExit(main())
