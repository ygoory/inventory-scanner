# התקנת Python וכל הדרישות למערכת ספירת מלאי
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

Write-Host "=== התקנת Python ===" -ForegroundColor Cyan

$pythonUrl = "https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe"
$installer = "$env:TEMP\python-installer.exe"

Write-Host "מוריד Python 3.11..." -ForegroundColor Yellow
Invoke-WebRequest -Uri $pythonUrl -OutFile $installer -UseBasicParsing

Write-Host "מתקין Python..." -ForegroundColor Yellow
Start-Process -FilePath $installer -ArgumentList "/quiet", "InstallAllUsers=1", "PrependPath=1", "Include_pip=1" -Wait

# רענן PATH
$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")

Write-Host "=== מתקין ספריות Python ===" -ForegroundColor Cyan
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

python -m pip install --upgrade pip
python -m pip install flask opencv-python pyzbar openpyxl Pillow numpy easyocr

Write-Host ""
Write-Host "=== ההתקנה הושלמה! ===" -ForegroundColor Green
Write-Host "להפעלת המערכת הרץ: הפעלה.ps1" -ForegroundColor Green
Read-Host "לחץ Enter לסיום"
