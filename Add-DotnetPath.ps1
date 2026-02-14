$dotnetPath = "D:\a03\dotnet"

# Check if the directory exists
if (-not (Test-Path $dotnetPath)) {
    Write-Host "Error: Directory '$dotnetPath' does not exist." -ForegroundColor Red
    exit 1
}

# Get current User Path
$currentPath = [Environment]::GetEnvironmentVariable("Path", "User")

# Check if path is already present
if ($currentPath -like "*$dotnetPath*") {
    Write-Host "The path '$dotnetPath' is already in the User environment variables." -ForegroundColor Yellow
}
else {
    # Add to Path
    $newPath = "$currentPath;$dotnetPath"
    [Environment]::SetEnvironmentVariable("Path", $newPath, "User")
    Write-Host "Successfully added '$dotnetPath' to User environment variables (Path)." -ForegroundColor Green
}

# Check and set DOTNET_ROOT
$currentDotnetRoot = [Environment]::GetEnvironmentVariable("DOTNET_ROOT", "User")
if ($currentDotnetRoot -ne $dotnetPath) {
    [Environment]::SetEnvironmentVariable("DOTNET_ROOT", $dotnetPath, "User")
    Write-Host "Successfully set DOTNET_ROOT to '$dotnetPath'." -ForegroundColor Green
} else {
    Write-Host "DOTNET_ROOT is already set correctly." -ForegroundColor Yellow
}

Write-Host "You may need to restart your terminal or computer for changes to take effect." -ForegroundColor Cyan

# Also attempt to set for current session for immediate testing (though this only affects this process)
$env:Path += ";$dotnetPath"
$env:DOTNET_ROOT = $dotnetPath

# Verify dotnet command
try {
    $version = & "$dotnetPath\dotnet.exe" --version
    Write-Host "Dotnet version: $version" -ForegroundColor Green
}
catch {
    Write-Host "Could not run dotnet command immediately." -ForegroundColor Yellow
}
