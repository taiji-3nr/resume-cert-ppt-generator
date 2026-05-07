$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$venv = Join-Path $root '.venv'

function Find-Python {
  $candidates = @('py', 'python', 'python3')
  foreach ($cmd in $candidates) {
    $found = Get-Command $cmd -ErrorAction SilentlyContinue
    if ($found) { return $cmd }
  }
  return $null
}

$python = Find-Python
if (-not $python) {
  Write-Error @'
Python was not found.

Install Python 3.11 or later from:
https://www.python.org/downloads/windows/

During installation, enable "Add python.exe to PATH", then run this script again.
'@
}

if (-not (Test-Path $venv)) {
  if ($python -eq 'py') {
    py -3 -m venv $venv
  } else {
    & $python -m venv $venv
  }
}

$venvPython = Join-Path $venv 'Scripts\python.exe'
& $venvPython -m pip install --upgrade pip
& $venvPython -m pip install -r (Join-Path $root 'requirements.txt')
& $venvPython -m pip install -e $root

Write-Host ''
Write-Host 'Environment is ready.'
Write-Host "Run: $venvPython -m resume_cert_ppt.generate_ppt"
