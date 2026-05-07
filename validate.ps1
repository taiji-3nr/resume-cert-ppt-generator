$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$dataPath = Join-Path $root 'data\certifications.json'
$pptPath = Join-Path $root 'out\資格取得ヒストリー_太地稔_20260103.pptx'
$pythonPptPath = Join-Path $root 'out\資格取得ヒストリー_太地稔_20260103_python.pptx'

$data = Get-Content -LiteralPath $dataPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ($data.certifications.Count -ne 14) {
  throw "Expected 14 certifications, got $($data.certifications.Count)."
}

foreach ($script in @('setup.ps1', 'build_cert_history_ppt.ps1', 'validate.ps1')) {
  $errors = $null
  $tokens = $null
  [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $root $script), [ref]$tokens, [ref]$errors) | Out-Null
  if ($errors) {
    throw "$script has syntax errors: $($errors -join '; ')"
  }
}

Add-Type -AssemblyName System.IO.Compression.FileSystem

$imagePath = Join-Path $root 'out\assets\certification-history-bg.png'
$expectImageMedia = Test-Path $imagePath

function Test-Pptx($path, $label, $expectMedia) {
  if (-not (Test-Path $path)) {
    throw "PowerPoint output was not found: $path"
  }

  $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
  try {
    $slides = ($zip.Entries | Where-Object { $_.FullName -like 'ppt/slides/slide*.xml' }).Count
    $media = ($zip.Entries | Where-Object { $_.FullName -like 'ppt/media/*' }).Count
    if ($slides -ne 5) { throw "$label expected 5 slides, got $slides." }
    if ($expectMedia -and $media -lt 1) { throw "$label expected at least 1 embedded media file, got $media." }

    foreach ($entry in $zip.Entries | Where-Object { $_.FullName -like '*.xml' -or $_.FullName -like '*.rels' }) {
      $reader = [System.IO.StreamReader]::new($entry.Open())
      try {
        [xml]$null = $reader.ReadToEnd()
      } finally {
        $reader.Dispose()
      }
    }
  } finally {
    $zip.Dispose()
  }
}

Test-Pptx $pptPath 'OpenXML PowerShell output' $expectImageMedia
Test-Pptx $pythonPptPath 'Python output' $expectImageMedia

Write-Host 'Validation OK'
Write-Host "certifications=$($data.certifications.Count)"
$mediaRule = if ($expectImageMedia) { 'media>=1' } else { 'media>=0 (no background image)' }
Write-Host "pptx outputs=2 slides=5 $mediaRule"
