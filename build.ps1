# Chạy test rồi đóng gói AutoApp.exe bằng môi trường Conda p310.
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

conda run -n p310 python -m pytest -q tests
if ($LASTEXITCODE -ne 0) { throw "Test lỗi, dừng build." }

conda run -n p310 pyinstaller AutoApp.spec --noconfirm
if ($LASTEXITCODE -ne 0) { throw "PyInstaller lỗi." }

Get-Item dist\AutoApp.exe | Select-Object Name, Length, LastWriteTime
