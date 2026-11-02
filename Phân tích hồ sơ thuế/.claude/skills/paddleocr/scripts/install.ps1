# Cai PaddleOCR MCP vao venv RIENG, khong dung .venv Javis.
$ErrorActionPreference = "Stop"
$Root = $PSScriptRoot
for ($i = 0; $i -lt 8; $i++) {
  if (Test-Path (Join-Path $Root "server\main.py")) { break }
  $Root = Split-Path -Parent $Root
}
if (-not (Test-Path (Join-Path $Root "server\main.py"))) {
  $Root = "D:\SECOND BRAIN\javis-os"
}
$Venv = Join-Path $Root ".venv-paddleocr"
$Py = Join-Path $Venv "Scripts\python.exe"
Write-Host "Root=$Root"
if (-not (Test-Path $Py)) {
  py -3.12 -m venv $Venv
}
& $Py -m pip install -U pip
& $Py -m pip install -U "paddleocr-mcp[local-cpu]"
& $Py -m paddleocr_mcp --help
Write-Host "Xong. Bat the PaddleOCR o trang Ket noi."
