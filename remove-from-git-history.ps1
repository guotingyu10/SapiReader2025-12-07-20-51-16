param(
    [Parameter(Mandatory = $true)]
    [string]$Target,

    [string]$RepoDir = "",

    [switch]$DryRun,

    [switch]$Force
)

$ErrorActionPreference = "Stop"

function Write-Step { param($Num, $Msg) Write-Host ""; Write-Host "========== [步骤 $Num] $Msg ==========" -ForegroundColor Cyan }
function Write-OK   { param($Msg) Write-Host "[OK] $Msg" -ForegroundColor Green }
function Write-Warn { param($Msg) Write-Host "[警告] $Msg" -ForegroundColor Yellow }
function Write-Err  { param($Msg) Write-Host "[错误] $Msg" -ForegroundColor Red }

Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "  git-filter-repo 历史清除脚本"
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "  目标路径:   $Target"
Write-Host "  模式:       $(if ($DryRun) { 'DRY-RUN (仅演示,不会真正修改历史)' } else { '实运行' })"
Write-Host "========================================================"

# ==================== 步骤 1: 解析仓库目录 ====================
Write-Step 1 "解析仓库目录"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

if ([string]::IsNullOrWhiteSpace($RepoDir)) {
    $startDir = (Get-Location).Path
} else {
    $startDir = (Resolve-Path -LiteralPath $RepoDir -ErrorAction Stop).Path
}

$RepoDir = $null
$current = $startDir
while ($current -and -not [string]::IsNullOrEmpty($current)) {
    if (Test-Path (Join-Path $current ".git")) {
        $RepoDir = $current
        break
    }
    $parent = Split-Path -Parent $current
    if ($parent -eq $current) { break }
    $current = $parent
}

if (-not $RepoDir) {
    $RepoDir = $startDir
}

if (-not (Test-Path (Join-Path $RepoDir ".git"))) {
    Write-Err "目录 $startDir 及上级目录中未找到 .git 仓库"
    exit 1
}
Write-OK "仓库根目录: $RepoDir"

Set-Location $RepoDir

# ==================== 步骤 2: 环境检查 ====================
Write-Step 2 "环境检查"

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    Write-Err "未找到 git 命令"
    exit 1
}
Write-OK "git 命令可用"

if (-not (Get-Command git-filter-repo -ErrorAction SilentlyContinue)) {
    Write-Err "未找到 git-filter-repo 命令"
    Write-Host "  安装方式（任选其一）:" -ForegroundColor Yellow
    Write-Host "    pip install git-filter-repo" -ForegroundColor Yellow
    Write-Host "    scoop install git-filter-repo" -ForegroundColor Yellow
    Write-Host "    choco install git-filter-repo" -ForegroundColor Yellow
    exit 1
}
Write-OK "git-filter-repo 命令可用"

# ==================== 步骤 3: 规范化目标路径 & 判断类型 ====================
Write-Step 3 "规范化目标路径"

$TargetRaw = $Target.Trim()
$TargetRaw = $TargetRaw.Replace('\', '/')

# 若传入的是绝对路径（含盘符），则转成相对于 RepoDir 的相对路径
$repoDirForMatch = $RepoDir.Replace('\', '/').TrimEnd('/')
if ($TargetRaw -match '^[A-Za-z]:/' -or $TargetRaw -match '^//' -or $TargetRaw -match '^\\\\') {
    $resolvedAbs = try { [System.IO.Path]::GetFullPath($TargetRaw.Replace('/', '\')).Replace('\', '/') } catch { $TargetRaw }
    if ($resolvedAbs.StartsWith($repoDirForMatch + '/', [System.StringComparison]::OrdinalIgnoreCase)) {
        $TargetClean = $resolvedAbs.Substring($repoDirForMatch.Length + 1)
        Write-Host "  绝对路径已转成相对路径: $TargetClean" -ForegroundColor Gray
    } else {
        Write-Err "绝对路径 '$TargetRaw' 不在仓库 '$RepoDir' 内，无法清除"
        exit 1
    }
} else {
    $TargetClean = $TargetRaw.TrimStart('/')
}

if ([string]::IsNullOrWhiteSpace($TargetClean)) {
    Write-Err "目标路径为空"
    exit 1
}

$TargetAbsPath = Join-Path $RepoDir $TargetClean.Replace('/', '\')

$isFolder = $false
$isFile   = $false

if (Test-Path -LiteralPath $TargetAbsPath -PathType Container) {
    $isFolder = $true
    Write-OK "目标识别为文件夹: $TargetClean"
} elseif (Test-Path -LiteralPath $TargetAbsPath -PathType Leaf) {
    $isFile = $true
    Write-OK "目标识别为文件: $TargetClean"
} else {
    Write-Warn "工作区中未找到 '$TargetClean',将按路径字面量在历史中尝试清除"
    if ($TargetClean.EndsWith('/') -or -not [System.IO.Path]::HasExtension($TargetClean)) {
        $isFolder = $true
        Write-Host "  推断为: 文件夹" -ForegroundColor Gray
    } else {
        $isFile = $true
        Write-Host "  推断为: 文件" -ForegroundColor Gray
    }
}

# ==================== 步骤 4: 备份提醒 & 确认 ====================
Write-Step 4 "备份提醒 & 二次确认"

$branch = git rev-parse --abbrev-ref HEAD 2>$null
$headLogCount = (git log --oneline -1 2>$null | Measure-Object).Count

Write-Host "  当前分支: $branch" -ForegroundColor Gray
Write-Host "  远程列表:" -ForegroundColor Gray
git remote -v

Write-Warn "此操作会彻底改写 git 历史（改变 commit hash）。"
Write-Warn "请确认已做好完整备份（例如复制整个 .git 目录到安全位置）。"
Write-Warn "执行后,所有协作者必须重新克隆仓库,否则会再次污染历史。"

if (-not $Force -and -not $DryRun) {
    $ans = Read-Host "确认从历史中彻底清除 '$TargetClean' ? (输入 YES 继续,其他任意输入退出)"
    if ($ans -ne "YES") {
        Write-Host "用户已取消操作"
        exit 0
    }
}

# ==================== 步骤 5: 备份目标路径（防止 filter-repo 从工作区删除文件） ====================
Write-Step 5 "备份目标路径到临时目录（filter-repo 执行后自动恢复）"

$tempRoot = [System.IO.Path]::GetTempPath()
$tempDirName = "git-filter-restore-$(Split-Path -Leaf $TargetClean)-$([guid]::NewGuid().ToString('N').Substring(0,8))"
$tempRestorePath = Join-Path $tempRoot $tempDirName
$targetExistsOnDisk = Test-Path -LiteralPath $TargetAbsPath

if ($targetExistsOnDisk) {
    if ($isFolder) {
        Copy-Item -Path $TargetAbsPath -Destination $tempRestorePath -Recurse -Force
        Write-Host "  已备份文件夹 -> $tempRestorePath" -ForegroundColor Gray
    } else {
        $destDir = Split-Path -Parent $tempRestorePath
        if (-not (Test-Path $destDir)) { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }
        Copy-Item -Path $TargetAbsPath -Destination $tempRestorePath -Force
        Write-Host "  已备份文件 -> $tempRestorePath" -ForegroundColor Gray
    }
    Write-OK "目标路径已备份"
} else {
    Write-Warn "工作区中 '$TargetClean' 不存在，跳过备份（仅清除历史）"
    $tempRestorePath = $null
}

# ==================== 步骤 6: 剥离远程引用 ====================
Write-Step 6 "剥离远程引用（防止 git-filter-repo 因远程存在而拒绝运行）"

$existingRemotes = git remote
$remotesSnapshot = @{}
foreach ($r in $existingRemotes) {
    $remotesSnapshot[$r] = git remote get-url $r
    git remote remove $r 2>&1 | Out-Null
    Write-Host "  已暂存远程: $r -> $($remotesSnapshot[$r])" -ForegroundColor Gray
}

# ==================== 步骤 7: 执行 git-filter-repo ====================
Write-Step 7 "执行 git-filter-repo"

$filterArgs = @()
if ($isFolder) {
    $filterArgs = @('--invert-paths', '--path', $TargetClean)
    Write-Host "  将从历史中清除文件夹: $TargetClean" -ForegroundColor Gray
} else {
    $filterArgs = @('--invert-paths', '--path', $TargetClean)
    Write-Host "  将从历史中清除文件: $TargetClean" -ForegroundColor Gray
}

if ($DryRun) {
    Write-Warn "DRY-RUN: 以下命令将被执行（不会真正运行）:"
    Write-Host "  git filter-repo $($filterArgs -join ' ') --force" -ForegroundColor Cyan
    if ($tempRestorePath) {
        Write-Host "  DRY-RUN 结束后，备份路径将保留: $tempRestorePath" -ForegroundColor Gray
    }
    Write-Warn "将远程引用恢复..."
    foreach ($r in $remotesSnapshot.Keys) {
        git remote add $r $remotesSnapshot[$r] 2>&1 | Out-Null
        Write-Host "  已恢复远程: $r -> $($remotesSnapshot[$r])" -ForegroundColor Gray
    }
    Write-OK "DRY-RUN 完成"
    exit 0
}

try {
    & git filter-repo @filterArgs --force
    if ($LASTEXITCODE -ne 0) {
        throw "git-filter-repo 退出码: $LASTEXITCODE"
    }
    Write-OK "历史改写完成"
} catch {
    Write-Err "git-filter-repo 执行失败: $_"
    Write-Warn "尝试恢复远程引用..."
    foreach ($r in $remotesSnapshot.Keys) {
        git remote add $r $remotesSnapshot[$r] 2>&1 | Out-Null
    }
    if ($tempRestorePath -and (Test-Path $tempRestorePath)) {
        try { Remove-Item -Path $tempRestorePath -Recurse -Force -ErrorAction SilentlyContinue } catch {}
    }
    exit 1
}

# ==================== 步骤 8: 恢复目标路径到工作区 + 加入 .gitignore + 停止跟踪 ====================
Write-Step 8 "恢复目标路径到工作区并加入 .gitignore"

if ($tempRestorePath -and (Test-Path -LiteralPath $tempRestorePath)) {
    $restoreDest = Join-Path $RepoDir $TargetClean
    $restoreParent = Split-Path -Parent $restoreDest
    if (-not (Test-Path -LiteralPath $restoreParent)) {
        New-Item -ItemType Directory -Path $restoreParent -Force | Out-Null
    }
    if (Test-Path -LiteralPath $restoreDest) {
        try { Remove-Item -Path $restoreDest -Recurse -Force -ErrorAction Stop } catch {}
    }
    Copy-Item -Path $tempRestorePath -Destination $restoreDest -Recurse -Force
    Write-OK "目标路径已恢复到工作区: $TargetClean"
    try { Remove-Item -Path $tempRestorePath -Recurse -Force -ErrorAction SilentlyContinue } catch {}
} else {
    Write-Host "  无备份数据需要恢复" -ForegroundColor Gray
}

$gitignorePath = Join-Path $RepoDir ".gitignore"
$gitignoreEntry = $TargetClean
if ($isFolder -and -not $gitignoreEntry.EndsWith('/')) {
    $gitignoreEntry = $gitignoreEntry + '/'
}

$cleanedContent = @()
$foundCorrect = $false
$hadBadEntries = $false
if (Test-Path -LiteralPath $gitignorePath) {
    $existingLines = Get-Content -Path $gitignorePath -Encoding UTF8
    foreach ($line in $existingLines) {
        $t = $line.Trim()
        if ($t -eq $gitignoreEntry) {
            $foundCorrect = $true
            $cleanedContent += $line
        } elseif ($t -match '^[A-Za-z]:' -or $t -match '^//' -or $t.StartsWith('/')) {
            $hadBadEntries = $true
        } else {
            $cleanedContent += $line
        }
    }
    $cleanedContent | Out-File -FilePath $gitignorePath -Encoding utf8
}

if ($hadBadEntries) {
    Write-Host "  已从 .gitignore 中移除含绝对路径的错误条目" -ForegroundColor Gray
}

if (-not $foundCorrect) {
    Add-Content -Path $gitignorePath -Value "" -Encoding UTF8
    Add-Content -Path $gitignorePath -Value "# Added by remove-from-git-history.ps1" -Encoding UTF8
    Add-Content -Path $gitignorePath -Value $gitignoreEntry -Encoding UTF8
    Write-OK "已将 '$gitignoreEntry' 加入 .gitignore"
} else {
    Write-Host "  '$gitignoreEntry' 已在 .gitignore 中" -ForegroundColor Gray
}

Set-Location $RepoDir
try {
    if ($isFolder) {
        & git rm -r --cached --ignore-unmatch $TargetClean 2>&1 | Out-Null
    } else {
        & git rm --cached --ignore-unmatch $TargetClean 2>&1 | Out-Null
    }
    if ($LASTEXITCODE -eq 0) {
        Write-OK "已从 git 索引中移除 '$TargetClean'（不再跟踪，工作区文件保留）"
    } else {
        Write-Host "  '$TargetClean' 当前不在 git 索引中" -ForegroundColor Gray
    }
} catch {
    Write-Host "  git rm --cached 跳过: $_" -ForegroundColor Gray
}

# ==================== 步骤 9: 恢复远程 & 提示后续操作 ====================
Write-Step 9 "恢复远程 & 提示后续操作"

foreach ($r in $remotesSnapshot.Keys) {
    git remote add $r $remotesSnapshot[$r] 2>&1 | Out-Null
    Write-Host "  已恢复远程: $r -> $($remotesSnapshot[$r])" -ForegroundColor Gray
}

Write-Host ""
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "  本地历史清除完成!"
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "  请执行后续操作:" -ForegroundColor Yellow
Write-Host "    1) 本地验证: git log --all --full-history --oneline -- '$TargetClean'" -ForegroundColor Yellow
Write-Host "    2) 强制推送: git push -f --all origin" -ForegroundColor Yellow
Write-Host "    3) 若有 tag:  git push -f --tags origin" -ForegroundColor Yellow
Write-Host "    4) 通知协作者: 需重新克隆,或执行 git reset --hard origin/<branch>" -ForegroundColor Yellow
Write-Host "========================================================" -ForegroundColor Cyan
