<#
.SYNOPSIS
    Creates a desktop shortcut for DriveTestApp with hidden console (pythonw).
#>
$ErrorActionPreference = "Stop"

try {
    $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    $ProjectRoot = Split-Path -Parent $ScriptDir
    $TargetFile = Join-Path $ProjectRoot "run.py"

    # Search for desktop path (supporting OneDrive redirection)
    $DesktopPath = [Environment]::GetFolderPath("Desktop")
    if ($env:OneDrive -and (Test-Path (Join-Path $env:OneDrive "Bureau"))) {
        $DesktopPath = Join-Path $env:OneDrive "Bureau"
    } elseif ($env:OneDrive -and (Test-Path (Join-Path $env:OneDrive "Desktop"))) {
        $DesktopPath = Join-Path $env:OneDrive "Desktop"
    }

    $ShortcutPath = Join-Path $DesktopPath "DriveTestApp.lnk"
    
    # Use pythonw.exe to run without a black console window
    $PythonExe = (Get-Command pythonw.exe -ErrorAction SilentlyContinue).Source
    if (-not $PythonExe) {
        $PythonExe = (Get-Command python.exe).Source
    }

    $WshShell = New-Object -comObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($ShortcutPath)
    $Shortcut.TargetPath = $PythonExe
    $Shortcut.Arguments = "`"$TargetFile`""
    $Shortcut.WorkingDirectory = $ProjectRoot
    $Shortcut.Description = "DriveTestApp - Network Automation & RNP Processing"
    $Shortcut.Save()
    
    Write-Host "Modern shortcut created successfully at: $ShortcutPath" -ForegroundColor Green
}
catch {
    Write-Host "Error creating shortcut: $_" -ForegroundColor Red
    exit 1
}
