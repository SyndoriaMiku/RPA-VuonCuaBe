# Chạy test rồi đóng gói AutoApp.exe bằng môi trường Conda p310.
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

conda run -n p310 python -m pytest -q tests
if ($LASTEXITCODE -ne 0) { throw "Test lỗi, dừng build." }

# Ghi mã commit + ngày + dấu -dirty vào version.py để bản build tự báo phiên bản.
# Luôn trả version.py về "dev" sau khi build, kể cả khi PyInstaller lỗi.
$versionFile = Join-Path $PSScriptRoot "version.py"
$originalVersion = [IO.File]::ReadAllText($versionFile)
$commit = (git rev-parse --short HEAD).Trim()
$dirty = if (git status --porcelain) { "-dirty" } else { "" }
$buildId = "$commit $(Get-Date -Format 'yyyy-MM-dd')$dirty"
$utf8NoBom = New-Object System.Text.UTF8Encoding $false
$patchedVersion = $originalVersion -replace 'BUILD = ".*"', ('BUILD = "' + $buildId + '"')
[IO.File]::WriteAllText($versionFile, $patchedVersion, $utf8NoBom)
Write-Output "Phiên bản build: $buildId"

try {
    conda run -n p310 pyinstaller AutoApp.spec --noconfirm
    if ($LASTEXITCODE -ne 0) { throw "PyInstaller lỗi." }
}
finally {
    [IO.File]::WriteAllText($versionFile, $originalVersion, $utf8NoBom)
}

Get-Item dist\AutoApp.exe | Select-Object Name, Length, LastWriteTime
