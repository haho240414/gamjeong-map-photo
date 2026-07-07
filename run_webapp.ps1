$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$VenvPython = Join-Path $ScriptDir ".venv\Scripts\python.exe"
$LocalPython = "C:\Users\haho1\AppData\Local\Python\bin\python.exe"

if (Test-Path $VenvPython) {
  $Python = $VenvPython
} elseif (Test-Path $LocalPython) {
  $Python = $LocalPython
} else {
  $Python = "python"
}

& $Python -X utf8 (Join-Path $ScriptDir "webapp_server.py") --host 127.0.0.1 --port 8877 --open
