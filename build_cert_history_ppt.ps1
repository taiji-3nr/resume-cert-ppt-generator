$ErrorActionPreference = 'Stop'

$workspace = 'E:\AI\Codex\動画編集'
$outDir = Join-Path $workspace 'out'
$assetDir = Join-Path $outDir 'assets'
$workDir = Join-Path $outDir 'pptx_build'
$pptxPath = Join-Path $outDir '資格取得ヒストリー_太地稔_20260103.pptx'
$sourceImage = 'C:\Users\taiji\.codex\generated_images\019e02d4-f008-7d40-aae0-e8fb8e0655ce\ig_0e33497e10a7c40c0169fca194244c8191ae54577403c19933.png'
$projectImage = Join-Path $assetDir 'certification-history-bg.png'

if (!(Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }
if (!(Test-Path $assetDir)) { New-Item -ItemType Directory -Path $assetDir | Out-Null }
Copy-Item -LiteralPath $sourceImage -Destination $projectImage -Force

if (Test-Path $workDir) { Remove-Item -LiteralPath $workDir -Recurse -Force }
New-Item -ItemType Directory -Path $workDir | Out-Null

function New-Dir($path) {
  if (!(Test-Path $path)) { New-Item -ItemType Directory -Path $path | Out-Null }
}

function Write-Utf8($path, $content) {
  $parent = Split-Path $path -Parent
  New-Dir $parent
  [System.IO.File]::WriteAllText($path, $content, [System.Text.UTF8Encoding]::new($false))
}

function X($text) {
  return [System.Security.SecurityElement]::Escape([string]$text)
}

function Shape($id, $name, $x, $y, $cx, $cy, $fill, $line, $text, $fontSize, $color, $bold = $false, $align = 'ctr') {
  $b = if ($bold) { ' b="1"' } else { '' }
  $ln = if ($line) { "<a:ln w=`"12700`"><a:solidFill><a:srgbClr val=`"$line`"/></a:solidFill></a:ln>" } else { '<a:ln><a:noFill/></a:ln>' }
  $fillXml = if ($fill) { "<a:solidFill><a:srgbClr val=`"$fill`"/></a:solidFill>" } else { '<a:noFill/>' }
  $paras = ''
  foreach ($part in ([string]$text).Split("`n")) {
    $paras += "<a:p><a:pPr algn=`"$align`"/><a:r><a:rPr lang=`"ja-JP`" sz=`"$fontSize`"$b><a:solidFill><a:srgbClr val=`"$color`"/></a:solidFill><a:latin typeface=`"Yu Gothic`"/><a:ea typeface=`"Yu Gothic`"/></a:rPr><a:t>$(X $part)</a:t></a:r></a:p>"
  }
  return @"
<p:sp>
  <p:nvSpPr><p:cNvPr id="$id" name="$(X $name)"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="$x" y="$y"/><a:ext cx="$cx" cy="$cy"/></a:xfrm>
    <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
    $fillXml
    $ln
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" anchor="mid"><a:spAutoFit/></a:bodyPr><a:lstStyle/>
    $paras
  </p:txBody>
</p:sp>
"@
}

function TextBox($id, $name, $x, $y, $cx, $cy, $text, $fontSize, $color, $bold = $false, $align = 'l') {
  return Shape $id $name $x $y $cx $cy $null $null $text $fontSize $color $bold $align
}

function Picture($id, $name, $rid, $x, $y, $cx, $cy) {
  return @"
<p:pic>
  <p:nvPicPr><p:cNvPr id="$id" name="$(X $name)"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
  <p:blipFill><a:blip r:embed="$rid"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
  <p:spPr><a:xfrm><a:off x="$x" y="$y"/><a:ext cx="$cx" cy="$cy"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
</p:pic>
"@
}

function SlideXml($body) {
  return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    $body
  </p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>
"@
}

function SlideRels($hasImage) {
  $img = if ($hasImage) { '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/certification-history-bg.png"/>' } else { '' }
  return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  $img
</Relationships>
"@
}

$W = 12192000
$H = 6858000

$slides = @()
$slides += SlideXml ((Picture 2 'generated certification history background' 'rId2' 0 0 $W $H) +
  (Shape 3 'title panel' 610000 560000 6600000 1850000 'FFFFFF' $null '資格取得ヒストリー' 4200 '102235' $true 'l') +
  (TextBox 4 'subtitle' 760000 1640000 6100000 650000 '太地 稔 | 1978 - 2022' 2200 '2B5D70' $true 'l') +
  (Shape 5 'summary' 760000 2520000 5300000 1050000 '0F2235' '39A5A9' '情報処理から教育・Office・認定インストラクターへ、実務と指導をつなぐ資格の軌跡' 1700 'FFFFFF' $false 'l'))

$body2 = TextBox 2 'title' 560000 280000 9500000 520000 '基礎技術を固めた時期' 3200 '102235' $true 'l'
$body2 += Shape 3 'line' 860000 2500000 10400000 65000 '39A5A9' $null '' 100 'FFFFFF'
$events1 = @(
  @{x=850000; y=1350000; yr='1978.02'; t="第二種情報処理技術者`n(現: 基本情報処理技術者)"},
  @{x=3650000; y=3150000; yr='1995.12'; t="初級システムアドミニストレータ`n(現: ITパスポート)"},
  @{x=6750000; y=1350000; yr='1997.07'; t='データベーススペシャリスト'}
)
$sid = 4
foreach ($e in $events1) {
  $body2 += Shape $sid "year" $e.x $e.y 1850000 520000 '102235' $null $e.yr 2000 'FFFFFF' $true 'ctr'; $sid++
  $body2 += Shape $sid "event" $e.x ($e.y + 610000) 2500000 760000 'FFFFFF' 'D9E3EA' $e.t 1450 '102235' $false 'ctr'; $sid++
}
$body2 += TextBox $sid 'note' 560000 5800000 9800000 500000 'ITキャリア初期から中堅期にかけて、業務システム開発の土台となる国家資格・専門資格を取得。' 1500 '486575' $false 'l'
$slides += SlideXml $body2

$body3 = TextBox 2 'title' 560000 280000 9800000 520000 '教育・指導領域への展開' 3200 '102235' $true 'l'
$body3 += Shape 3 'phase1' 640000 1220000 2500000 3750000 'F4F8FA' 'D5E3EA' 'Office活用' 1900 '102235' $true 'ctr'
$body3 += Shape 4 'phase2' 3560000 1220000 2500000 3750000 'F8F5EC' 'E2D7B8' 'ICT支援' 1900 '102235' $true 'ctr'
$body3 += Shape 5 'phase3' 6480000 1220000 2500000 3750000 'F1F7F4' 'C9DDD2' 'プログラミング教育' 1900 '102235' $true 'ctr'
$body3 += Shape 6 'phase4' 9400000 1220000 1900000 3750000 'F5F2F8' 'DAD0E6' '認定講師' 1900 '102235' $true 'ctr'
$body3 += TextBox 7 'p1' 900000 2000000 2000000 2100000 "2016 MOS 2013 Master`n2019 Excel 1級`n2020 Access 1級`n2022 PowerPoint上級`n2022 Excelビジネススキル" 1350 '243746' $false 'l'
$body3 += TextBox 8 'p2' 3820000 2000000 2000000 1550000 "2017 ICT支援員能力認定`n2020 P検3級" 1450 '243746' $false 'l'
$body3 += TextBox 9 'p3' 6740000 2000000 2000000 1550000 "2018 Scratch Gold`n2019 Scratch Silver" 1450 '243746' $false 'l'
$body3 += TextBox 10 'p4' 9620000 2000000 1500000 1550000 "2022 サーティファイ`n認定インストラクター" 1450 '243746' $false 'l'
$body3 += Shape 11 'message' 1150000 5550000 9800000 700000 '102235' $null 'セカンドキャリアでは、使える力を教える力へ転換。資格の幅がそのまま指導対象の広がりを示している。' 1550 'FFFFFF' $false 'ctr'
$slides += SlideXml $body3

$body4 = TextBox 2 'title' 560000 280000 9800000 520000 'スキル領域マップ' 3200 '102235' $true 'l'
$body4 += Shape 3 'center' 4660000 2500000 2900000 1000000 '102235' '39A5A9' '実務 × 教育' 2400 'FFFFFF' $true 'ctr'
$body4 += Shape 4 'a' 900000 1250000 3100000 1050000 'FFFFFF' '39A5A9' "IT基礎・業務理解`n第二種情報処理 / シスアド / P検" 1500 '102235' $false 'ctr'
$body4 += Shape 5 'b' 8000000 1250000 3100000 1050000 'FFFFFF' '39A5A9' "専門技術`nデータベーススペシャリスト" 1500 '102235' $false 'ctr'
$body4 += Shape 6 'c' 900000 4550000 3100000 1050000 'FFFFFF' 'C49A2C' "Office実務`nMOS / Excel / Access / PowerPoint" 1500 '102235' $false 'ctr'
$body4 += Shape 7 'd' 8000000 4550000 3100000 1050000 'FFFFFF' 'C49A2C' "指導・教育`nICT支援員 / Scratch / 認定インストラクター" 1500 '102235' $false 'ctr'
$body4 += TextBox 8 'note' 4300000 3850000 3700000 650000 '資格群は、40年のIT実務経験とセカンドキャリアの教育実践を接続する証跡。' 1450 '486575' $false 'ctr'
$slides += SlideXml $body4

$body5 = (Picture 2 'generated certification history background' 'rId2' 0 0 $W $H)
$body5 += Shape 3 'panel' 650000 620000 5800000 5200000 'FFFFFF' $null '資格取得のストーリー' 3300 '102235' $true 'l'
$body5 += TextBox 4 'story' 900000 1600000 5100000 2900000 "1. IT基礎を国家資格で体系化`n2. DBなど専門技術で開発マネジメントを補強`n3. MOS・サーティファイでOffice実務力を可視化`n4. ICT支援・Scratch・認定講師で教育領域へ展開" 1650 '243746' $false 'l'
$body5 += Shape 5 'closing' 900000 4880000 5000000 650000 '102235' '39A5A9' '「作る」経験を、「教える」価値へ。' 2000 'FFFFFF' $true 'ctr'
$slides += SlideXml $body5

New-Dir (Join-Path $workDir '_rels')
New-Dir (Join-Path $workDir 'docProps')
New-Dir (Join-Path $workDir 'ppt\_rels')
New-Dir (Join-Path $workDir 'ppt\slides\_rels')
New-Dir (Join-Path $workDir 'ppt\slideLayouts\_rels')
New-Dir (Join-Path $workDir 'ppt\slideMasters\_rels')
New-Dir (Join-Path $workDir 'ppt\theme')
New-Dir (Join-Path $workDir 'ppt\media')

Copy-Item -LiteralPath $projectImage -Destination (Join-Path $workDir 'ppt\media\certification-history-bg.png') -Force

$overrideSlides = ''
for ($i=1; $i -le $slides.Count; $i++) {
  Write-Utf8 (Join-Path $workDir "ppt\slides\slide$i.xml") $slides[$i-1]
  Write-Utf8 (Join-Path $workDir "ppt\slides\_rels\slide$i.xml.rels") (SlideRels ($i -eq 1 -or $i -eq 5))
  $overrideSlides += "<Override PartName=`"/ppt/slides/slide$i.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.presentationml.slide+xml`"/>"
}

Write-Utf8 (Join-Path $workDir '[Content_Types].xml') @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  $overrideSlides
</Types>
"@

Write-Utf8 (Join-Path $workDir '_rels\.rels') @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"@

$slideIdList = ''
$presRels = '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>'
for ($i=1; $i -le $slides.Count; $i++) {
  $rid = $i + 1
  $sid = 255 + $i
  $slideIdList += "<p:sldId id=`"$sid`" r:id=`"rId$rid`"/>"
  $presRels += "<Relationship Id=`"rId$rid`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide`" Target=`"slides/slide$i.xml`"/>"
}
Write-Utf8 (Join-Path $workDir 'ppt\presentation.xml') @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst>$slideIdList</p:sldIdLst>
  <p:sldSz cx="$W" cy="$H" type="wide"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>
"@

Write-Utf8 (Join-Path $workDir 'ppt\_rels\presentation.xml.rels') @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">$presRels</Relationships>
"@

Write-Utf8 (Join-Path $workDir 'ppt\slideMasters\slideMaster1.xml') @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
</p:sldMaster>
"@

Write-Utf8 (Join-Path $workDir 'ppt\slideMasters\_rels\slideMaster1.xml.rels') @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>
"@

Write-Utf8 (Join-Path $workDir 'ppt\slideLayouts\slideLayout1.xml') @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
  <p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
</p:sldLayout>
"@

Write-Utf8 (Join-Path $workDir 'ppt\slideLayouts\_rels\slideLayout1.xml.rels') @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>
"@

Write-Utf8 (Join-Path $workDir 'ppt\theme\theme1.xml') @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Certification History">
  <a:themeElements>
    <a:clrScheme name="Custom"><a:dk1><a:srgbClr val="102235"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="243746"/></a:dk2><a:lt2><a:srgbClr val="F4F8FA"/></a:lt2><a:accent1><a:srgbClr val="39A5A9"/></a:accent1><a:accent2><a:srgbClr val="C49A2C"/></a:accent2><a:accent3><a:srgbClr val="486575"/></a:accent3><a:accent4><a:srgbClr val="D9E3EA"/></a:accent4><a:accent5><a:srgbClr val="F1F7F4"/></a:accent5><a:accent6><a:srgbClr val="F5F2F8"/></a:accent6><a:hlink><a:srgbClr val="2B5D70"/></a:hlink><a:folHlink><a:srgbClr val="5A4D72"/></a:folHlink></a:clrScheme>
    <a:fontScheme name="Yu Gothic"><a:majorFont><a:latin typeface="Yu Gothic"/><a:ea typeface="Yu Gothic"/></a:majorFont><a:minorFont><a:latin typeface="Yu Gothic"/><a:ea typeface="Yu Gothic"/></a:minorFont></a:fontScheme>
    <a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>
  </a:themeElements>
</a:theme>
"@

$now = (Get-Date).ToUniversalTime().ToString('s') + 'Z'
Write-Utf8 (Join-Path $workDir 'docProps\core.xml') @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>資格取得ヒストリー_太地稔</dc:title><dc:creator>Codex</dc:creator><cp:lastModifiedBy>Codex</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">$now</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">$now</dcterms:modified>
</cp:coreProperties>
"@

Write-Utf8 (Join-Path $workDir 'docProps\app.xml') @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Codex</Application><PresentationFormat>On-screen Show (16:9)</PresentationFormat><Slides>$($slides.Count)</Slides></Properties>
"@

if (Test-Path $pptxPath) { Remove-Item -LiteralPath $pptxPath -Force }
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$fs = [System.IO.File]::Open($pptxPath, [System.IO.FileMode]::CreateNew)
$archive = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
try {
  $base = (Resolve-Path $workDir).Path.TrimEnd('\')
  foreach ($file in Get-ChildItem -LiteralPath $workDir -Recurse -File) {
    $entryName = $file.FullName.Substring($base.Length + 1).Replace('\', '/')
    $entry = $archive.CreateEntry($entryName, [System.IO.Compression.CompressionLevel]::Optimal)
    $inStream = [System.IO.File]::OpenRead($file.FullName)
    try {
      $outStream = $entry.Open()
      try { $inStream.CopyTo($outStream) } finally { $outStream.Dispose() }
    } finally {
      $inStream.Dispose()
    }
  }
} finally {
  $archive.Dispose()
  $fs.Dispose()
}

Get-Item -LiteralPath $pptxPath | Select-Object FullName,Length,LastWriteTime
