@echo off
rem 图片转PPTX v3.0.3
setlocal
set "IMG2PPTX_SCRIPT=%~f0"
set "IMG2PPTX_DIR=%~dp0"
set "IMG2PPTX_INPUT=%~1"
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -Command "$script=[IO.File]::ReadAllText($env:IMG2PPTX_SCRIPT,[Text.Encoding]::UTF8); $parts=[regex]::Split($script,'(?m)^# POWERSHELL_START\r?$',2); if($parts.Count -lt 2){throw 'PowerShell section missing'}; iex $parts[1]"
set "exitCode=%ERRORLEVEL%"
if not "%IMG2PPTX_NO_PAUSE%"=="1" pause
exit /b %exitCode%
# POWERSHELL_START
$ErrorActionPreference = "Stop"

[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

$SlideWidthEmu = 12192000
$SlideHeightEmu = 6858000
$Utf8NoBom = New-Object System.Text.UTF8Encoding $false

function Show-Message {
    param(
        [string]$Title,
        [string]$Message,
        [System.Windows.Forms.MessageBoxIcon]$Icon = [System.Windows.Forms.MessageBoxIcon]::Information
    )
    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        $Icon
    ) | Out-Null
}

function Fail {
    param([string]$Message)
    Write-Host ""
    Write-Host $Message
    Show-Message -Title "图片转PPTX失败" -Message $Message -Icon ([System.Windows.Forms.MessageBoxIcon]::Error)
    exit 1
}

function Write-Utf8 {
    param([string]$Path, [string]$Content)
    [System.IO.File]::WriteAllText($Path, $Content, $Utf8NoBom)
}

function Choose-InputFolder {
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "请选择包含数字结尾（至少两位）图片的文件夹"
    $dialog.ShowNewFolderButton = $false
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.SelectedPath
    }
    exit 0
}

function Get-UniqueOutputPath {
    param([string]$InputDir)

    $folderName = Split-Path -Path $InputDir -Leaf
    if ([string]::IsNullOrWhiteSpace($folderName)) {
        $folderName = "images"
    }

    $output = Join-Path $InputDir "$folderName.pptx"
    if (-not (Test-Path -LiteralPath $output)) {
        return $output
    }

    for ($i = 2; $i -lt 1000; $i++) {
        $candidate = Join-Path $InputDir ("{0}_{1}.pptx" -f $folderName, $i)
        if (-not (Test-Path -LiteralPath $candidate)) {
            return $candidate
        }
    }

    Fail "输出文件已存在，且无法生成新的文件名。"
}

function Scan-Images {
    param([string]$InputDir)

    $byNumber = @{}
    $numbers = New-Object System.Collections.Generic.List[string]
    $skipped = New-Object System.Collections.Generic.List[string]
    $extensions = @(".jpg", ".jpeg", ".png")

    foreach ($file in Get-ChildItem -LiteralPath $InputDir -File) {
        $extension = $file.Extension.ToLowerInvariant()
        if ($extensions -notcontains $extension) {
            continue
        }

        $stem = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        if ($stem -match "(?<!\d)(\d{2,})$") {
            $numberInt = [int64]$Matches[1]
            if ($numberInt -lt 1) {
                continue
            }

            $key = "{0:D3}" -f $numberInt
            if ($byNumber.ContainsKey($key)) {
                $first = $byNumber[$key].Name
                Fail "发现重复编号 $key：`n$first`n$($file.Name)`n`n例如 01 和 001 会被视为同一页，请保留其中一个文件后再运行。"
            }

            $byNumber[$key] = $file
            $numbers.Add($key) | Out-Null
        }
        else {
            $skipped.Add($file.Name) | Out-Null
        }
    }

    return [PSCustomObject]@{
        ByNumber = $byNumber
        Numbers = $numbers
        Skipped = $skipped
    }
}

function Get-ImagePlacement {
    param([System.IO.FileInfo]$File)

    $image = $null
    try {
        $image = [System.Drawing.Image]::FromFile($File.FullName)
        $width = [int64]$image.Width
        $height = [int64]$image.Height
        $orientation = 1
        if ($image.PropertyIdList -contains 0x0112) {
            $orientation = [int]$image.GetPropertyItem(0x0112).Value[0]
        }
        if ($orientation -ge 5 -and $orientation -le 8) {
            $swap = $width; $width = $height; $height = $swap
        }
    }
    finally {
        if ($image -ne $null) {
            $image.Dispose()
        }
    }

    if ($width -le 0 -or $height -le 0) {
        Fail "图片尺寸异常：$($File.Name)"
    }

    if ($width * $SlideHeightEmu -gt $height * $SlideWidthEmu) {
        $cx = [int64]$SlideWidthEmu
        $cy = [int64](($SlideWidthEmu * $height + [math]::Floor($width / 2)) / $width)
        $x = [int64]0
        $y = [int64](($SlideHeightEmu - $cy) / 2)
    }
    else {
        $cy = [int64]$SlideHeightEmu
        $cx = [int64](($SlideHeightEmu * $width + [math]::Floor($height / 2)) / $height)
        $x = [int64](($SlideWidthEmu - $cx) / 2)
        $y = [int64]0
    }

    return [PSCustomObject]@{
        X = $x
        Y = $y
        Cx = $cx
        Cy = $cy
    }
}

function Write-StaticParts {
    param([string]$Work, [int]$SlideCount, [string]$Title)

    $slideOverrides = (1..$SlideCount | ForEach-Object {
        "  <Override PartName=`"/ppt/slides/slide$_.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.presentationml.slide+xml`"/>"
    }) -join "`r`n"

    Write-Utf8 (Join-Path $Work "[Content_Types].xml") @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="jpg" ContentType="image/jpeg"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>
  <Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>
  <Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
$slideOverrides
</Types>
"@

    Write-Utf8 (Join-Path $Work "_rels\.rels") @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
'@

    $slideTitles = (1..$SlideCount | ForEach-Object {
        "    <vt:lpstr>Slide $_</vt:lpstr>"
    }) -join "`r`n"

    Write-Utf8 (Join-Path $Work "docProps\app.xml") @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>图片转PPTX-Windows.cmd</Application>
  <PresentationFormat>On-screen Show (16:9)</PresentationFormat>
  <Slides>$SlideCount</Slides>
  <Notes>0</Notes>
  <HiddenSlides>0</HiddenSlides>
  <MMClips>0</MMClips>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs>
    <vt:vector size="2" baseType="variant">
      <vt:variant><vt:lpstr>Slides</vt:lpstr></vt:variant>
      <vt:variant><vt:i4>$SlideCount</vt:i4></vt:variant>
    </vt:vector>
  </HeadingPairs>
  <TitlesOfParts>
    <vt:vector size="$SlideCount" baseType="lpstr">
$slideTitles
    </vt:vector>
  </TitlesOfParts>
  <Company/>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0000</AppVersion>
</Properties>
"@

    $now = [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")
    $titleXml = [System.Security.SecurityElement]::Escape($Title)
    Write-Utf8 (Join-Path $Work "docProps\core.xml") @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>$titleXml</dc:title>
  <dc:creator>图片转PPTX-Windows.cmd</dc:creator>
  <cp:lastModifiedBy>图片转PPTX-Windows.cmd</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">$now</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">$now</dcterms:modified>
</cp:coreProperties>
"@

    $slideIds = (1..$SlideCount | ForEach-Object {
        $id = 255 + $_
        $rid = $_ + 1
        "    <p:sldId id=`"$id`" r:id=`"rId$rid`"/>"
    }) -join "`r`n"

    Write-Utf8 (Join-Path $Work "ppt\presentation.xml") @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst>
$slideIds
  </p:sldIdLst>
  <p:sldSz cx="$SlideWidthEmu" cy="$SlideHeightEmu" type="wide"/>
  <p:notesSz cx="6858000" cy="9144000"/>
  <p:defaultTextStyle/>
</p:presentation>
"@

    $slideRels = (1..$SlideCount | ForEach-Object {
        $rid = $_ + 1
        "  <Relationship Id=`"rId$rid`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide`" Target=`"slides/slide$_.xml`"/>"
    }) -join "`r`n"
    $presPropsId = $SlideCount + 2
    $viewPropsId = $SlideCount + 3
    $tableStylesId = $SlideCount + 4

    Write-Utf8 (Join-Path $Work "ppt\_rels\presentation.xml.rels") @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
$slideRels
  <Relationship Id="rId$presPropsId" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>
  <Relationship Id="rId$viewPropsId" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>
  <Relationship Id="rId$tableStylesId" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>
</Relationships>
"@

    Write-Utf8 (Join-Path $Work "ppt\presProps.xml") @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>
'@

    Write-Utf8 (Join-Path $Work "ppt\viewProps.xml") @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr>
  <p:slideViewPr><p:cSldViewPr><p:cViewPr varScale="1"><p:scale><a:sx n="100" d="100"/><a:sy n="100" d="100"/></p:scale><p:origin x="0" y="0"/></p:cViewPr><p:guideLst/></p:cSldViewPr></p:slideViewPr>
  <p:notesTextViewPr><p:cViewPr><p:scale><a:sx n="100" d="100"/><a:sy n="100" d="100"/></p:scale><p:origin x="0" y="0"/></p:cViewPr></p:notesTextViewPr>
  <p:gridSpacing cx="72008" cy="72008"/>
</p:viewPr>
'@

    Write-Utf8 (Join-Path $Work "ppt\tableStyles.xml") @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>
'@
}

function Write-ThemeAndLayouts {
    param([string]$Work)

    Write-Utf8 (Join-Path $Work "ppt\slideMasters\slideMaster1.xml") @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
  <p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>
</p:sldMaster>
'@

    Write-Utf8 (Join-Path $Work "ppt\slideMasters\_rels\slideMaster1.xml.rels") @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>
'@

    Write-Utf8 (Join-Path $Work "ppt\slideLayouts\slideLayout1.xml") @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             type="blank" preserve="1">
  <p:cSld name="Blank">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>
'@

    Write-Utf8 (Join-Path $Work "ppt\slideLayouts\_rels\slideLayout1.xml.rels") @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>
'@

    Write-Utf8 (Join-Path $Work "ppt\theme\theme1.xml") @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F1F1F"/></a:dk2>
      <a:lt2><a:srgbClr val="F2F2F2"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Aptos"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>
        <a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>
        <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>
        <a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill>
        <a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>
'@
}

function Write-Slide {
    param(
        [string]$Work,
        [int]$SlideIndex,
        [string]$PackageName,
        [int64]$X,
        [int64]$Y,
        [int64]$Cx,
        [int64]$Cy
    )

    Write-Utf8 (Join-Path $Work "ppt\slides\slide$SlideIndex.xml") @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:pic>
        <p:nvPicPr>
          <p:cNvPr id="2" name="Picture $SlideIndex"/>
          <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
          <p:nvPr/>
        </p:nvPicPr>
        <p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
        <p:spPr>
          <a:xfrm><a:off x="$X" y="$Y"/><a:ext cx="$Cx" cy="$Cy"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
      </p:pic>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>
"@

    Write-Utf8 (Join-Path $Work "ppt\slides\_rels\slide$SlideIndex.xml.rels") @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/$PackageName"/>
</Relationships>
"@
}

try {
    $scriptDir = [System.IO.Path]::GetFullPath($env:IMG2PPTX_DIR)
    $inputDir = $env:IMG2PPTX_INPUT
    $usedScriptFolder = $false

    if ([string]::IsNullOrWhiteSpace($inputDir)) {
        $inputDir = $scriptDir
        $usedScriptFolder = $true
    }
    else {
        $inputDir = [System.IO.Path]::GetFullPath($inputDir)
    }

    if (-not (Test-Path -LiteralPath $inputDir -PathType Container)) {
        Fail "输入路径不是文件夹：$inputDir"
    }

    $scan = Scan-Images $inputDir
    if ($scan.Numbers.Count -eq 0 -and $usedScriptFolder) {
        $inputDir = Choose-InputFolder
        $scan = Scan-Images $inputDir
    }

    if ($scan.Numbers.Count -eq 0) {
        Fail "没有找到文件名以两位及以上数字结尾的 JPG/PNG 图片。"
    }

    $sortedNumbers = $scan.Numbers | Sort-Object { [int64]$_ }
    $slideCount = @($sortedNumbers).Count
    $output = Get-UniqueOutputPath $inputDir
    $work = Join-Path ([System.IO.Path]::GetTempPath()) ("images-to-pptx-" + [Guid]::NewGuid().ToString("N"))

    $dirs = @(
        "_rels",
        "docProps",
        "ppt\_rels",
        "ppt\slides\_rels",
        "ppt\media",
        "ppt\slideMasters\_rels",
        "ppt\slideLayouts\_rels",
        "ppt\theme"
    )

    foreach ($dir in $dirs) {
        New-Item -ItemType Directory -Path (Join-Path $work $dir) -Force | Out-Null
    }

    Write-Host "正在生成 PPTX..."
    Write-Host "图片文件夹：$inputDir"
    Write-Host "输出文件：$output"
    Write-Host ""
    Write-Host "幻灯片顺序："

    Write-StaticParts -Work $work -SlideCount $slideCount -Title ([System.IO.Path]::GetFileNameWithoutExtension($output))
    Write-ThemeAndLayouts -Work $work

    $slideIndex = 1
    foreach ($number in $sortedNumbers) {
        $source = $scan.ByNumber[$number]
        $extension = $source.Extension.TrimStart(".").ToLowerInvariant()
        $packageName = "image$slideIndex.$extension"
        $placement = Get-ImagePlacement $source
        Write-Host ("  {0:D3}. [{1}] {2}" -f $slideIndex, $number, $source.Name)
        Copy-Item -LiteralPath $source.FullName -Destination (Join-Path $work "ppt\media\$packageName")
        Write-Slide -Work $work -SlideIndex $slideIndex -PackageName $packageName -X $placement.X -Y $placement.Y -Cx $placement.Cx -Cy $placement.Cy
        $slideIndex++
    }

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    if (Test-Path -LiteralPath $output) {
        Remove-Item -LiteralPath $output -Force
    }
    [System.IO.Compression.ZipFile]::CreateFromDirectory(
        $work,
        $output,
        [System.IO.Compression.CompressionLevel]::Optimal,
        $false
    )

    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue

    Write-Host ""
    Write-Host "完成：$output"
    Write-Host "共 $slideCount 页，图片保持原始比例居中放置。"
    Start-Process -FilePath $output
    Show-Message -Title "图片转PPTX完成" -Message "已生成 $slideCount 页 PPTX，图片未拉伸。`n`n$output"
}
catch {
    try {
        if ($work -and (Test-Path -LiteralPath $work)) {
            Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    catch {}
    Fail $_.Exception.Message
}
