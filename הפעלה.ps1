# הפעלת מערכת ספירת מלאי
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

# מצא Python
$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
$pythonExe = (Get-Command python -ErrorAction SilentlyContinue).Source

if (-not $pythonExe) {
    Write-Host "Python לא נמצא. הרץ קודם את התקנה.ps1" -ForegroundColor Red
    Read-Host "לחץ Enter לסיום"
    exit 1
}

Write-Host "=== מערכת ספירת מלאי ===" -ForegroundColor Cyan
Write-Host "Python: $pythonExe" -ForegroundColor Gray
Write-Host ""
Write-Host "השרת עולה... פתח בדפדפן:" -ForegroundColor Yellow
Write-Host "  http://localhost:5000" -ForegroundColor Green
Write-Host "  או מהטלפון: http://<IP-של-המחשב>:5000" -ForegroundColor Green
Write-Host ""
Write-Host "(לעצירה לחץ Ctrl+C)" -ForegroundColor Gray

# הצג IP מקומי
$ip = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.IPAddress -notmatch "^127\." -and $_.PrefixOrigin -ne "WellKnown" } | Select -First 1).IPAddress
if ($ip) { Write-Host "  IP מקומי: http://${ip}:5000" -ForegroundColor Green }

Write-Host ""
& $pythonExe app.py
