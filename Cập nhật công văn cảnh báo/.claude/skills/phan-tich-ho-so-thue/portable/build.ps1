# Dong goi exe portable. Chay tu may da co .venv-paddleocr.
$ErrorActionPreference = "Stop"
$Here = $PSScriptRoot
$Skill = Split-Path $Here -Parent
$Engine = Join-Path $Skill "engine"
$Repo = $Here
for ($i = 0; $i -lt 8; $i++) {
  if (Test-Path (Join-Path $Repo "server\main.py")) { break }
  $Repo = Split-Path -Parent $Repo
}
$Py = Join-Path $Repo ".venv-paddleocr\Scripts\python.exe"
if (-not (Test-Path $Py)) { throw "Thieu $Py" }
& $Py -m pip install -q pyinstaller openpyxl
$Dist = Join-Path $Here "dist"
$Work = Join-Path $Here "build"
if (Test-Path $Dist) { Remove-Item $Dist -Recurse -Force }
if (Test-Path $Work) { Remove-Item $Work -Recurse -Force }
Push-Location $Engine
try {
  & $Py -m PyInstaller --noconfirm --clean --onedir `
    --name PhanTichHoSoThue `
    --distpath $Dist `
    --workpath $Work `
    --specpath $Here `
    --hidden-import openpyxl `
    --hidden-import openpyxl.cell._writer `
    --exclude-module paddle `
    --exclude-module paddleocr `
    --exclude-module pandas `
    --exclude-module numpy `
    --exclude-module scipy `
    --exclude-module torch `
    --exclude-module tensorflow `
    --exclude-module PIL `
    --exclude-module cv2 `
    --console `
    (Join-Path $Engine "cli.py")
} finally {
  Pop-Location
}
Copy-Item (Join-Path $Here "Chay.bat") (Join-Path $Dist "Chay.bat") -Force
Copy-Item (Join-Path $Here "DOC.txt") (Join-Path $Dist "DOC.txt") -Force
foreach ($doc in @("HUONG-DAN-SU-DUNG.md", "HUONG-DAN-DAY-DU.md", "JAVIS-TAO-SKILL.md")) {
  $p = Join-Path $Here $doc
  if (Test-Path $p) { Copy-Item $p (Join-Path $Dist $doc) -Force }
}
$Zip = Join-Path $Here "PhanTichHoSoThue-portable.zip"
if (Test-Path $Zip) { Remove-Item $Zip -Force }
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($Dist, $Zip)
Write-Host "XONG $Zip"
Write-Host "Copy file zip sang USB, giai nen, chay Chay.bat"
