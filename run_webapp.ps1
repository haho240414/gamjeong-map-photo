$ErrorActionPreference = "Stop"
$env:PYTHONUTF8 = "1"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ScriptDir
$VenvPython = Join-Path $ScriptDir ".venv\Scripts\python.exe"

if (-not (Test-Path $VenvPython)) {
  Write-Host "Creating virtual environment..."
  if (Get-Command py -ErrorAction SilentlyContinue) { & py -3 -m venv .venv } else { & python -m venv .venv }
  & $VenvPython -m pip install -q --upgrade pip
  & $VenvPython -m pip install -q -r requirements.txt
}
if (-not (Test-Path ".env")) { Copy-Item ".env.example" ".env" }

& $VenvPython -X utf8 (Join-Path $ScriptDir "webapp_server.py") --host 127.0.0.1 --port 8877 --open
