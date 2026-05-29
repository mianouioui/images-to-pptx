#!/bin/zsh
# 图片转PPTX v3.0.2
set -euo pipefail
chmod +x "$0" >/dev/null 2>&1 || true

SLIDE_WIDTH_EMU=12192000
SLIDE_HEIGHT_EMU=6858000
OSASCRIPT=/usr/bin/osascript
ZIP=/usr/bin/zip
OPEN=/usr/bin/open
MKTEMP=/usr/bin/mktemp
SIPS=/usr/bin/sips
AWK=/usr/bin/awk
PERL=/usr/bin/perl

fail() {
  local message="$1"
  echo "$message" >&2
  "$OSASCRIPT" \
    -e 'on run argv' \
    -e 'display alert "图片转PPTX失败" message (item 1 of argv)' \
    -e 'end run' \
    "$message" >/dev/null 2>&1 || true
  exit 1
}

xml_escape() {
  local s="$1"
  s="${s//&/&amp;}"
  s="${s//</&lt;}"
  s="${s//>/&gt;}"
  s="${s//\"/&quot;}"
  s="${s//\'/&apos;}"
  print -r -- "$s"
}

choose_input_folder() {
  "$OSASCRIPT" -e 'POSIX path of (choose folder with prompt "请选择包含数字结尾（至少两位）图片的文件夹")'
}

unique_output_path() {
  local input_dir="$1"
  local folder_name="${input_dir:t}"
  [[ -z "$folder_name" ]] && folder_name="images"

  local output="$input_dir/$folder_name.pptx"
  if [[ ! -e "$output" ]]; then
    print -r -- "$output"
    return
  fi

  local index=2
  while [[ $index -lt 1000 ]]; do
    local candidate="$input_dir/${folder_name}_${index}.pptx"
    if [[ ! -e "$candidate" ]]; then
      print -r -- "$candidate"
      return
    fi
    index=$((index + 1))
  done

  fail "输出文件已存在，且无法生成新的文件名。"
}

write_static_parts() {
  local work="$1"
  local slide_count="$2"
  local title now
  title="$(xml_escape "$3")"
  now="$(date -u +"%Y-%m-%dT%H:%M:%SZ")"

  {
    print -r -- '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    print -r -- '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    print -r -- '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    print -r -- '  <Default Extension="xml" ContentType="application/xml"/>'
    print -r -- '  <Default Extension="jpg" ContentType="image/jpeg"/>'
    print -r -- '  <Default Extension="jpeg" ContentType="image/jpeg"/>'
    print -r -- '  <Default Extension="png" ContentType="image/png"/>'
    print -r -- '  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
    print -r -- '  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
    print -r -- '  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
    print -r -- '  <Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>'
    print -r -- '  <Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>'
    print -r -- '  <Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>'
    print -r -- '  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
    print -r -- '  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
    print -r -- '  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
    local i
    for (( i = 1; i <= slide_count; i++ )); do
      print -r -- "  <Override PartName=\"/ppt/slides/slide${i}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>"
    done
    print -r -- '</Types>'
  } > "$work/[Content_Types].xml"

  cat > "$work/_rels/.rels" <<'XML'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
XML

  {
    print -r -- '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    print -r -- '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"'
    print -r -- '            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
    print -r -- '  <Application>图片转PPTX.command</Application>'
    print -r -- '  <PresentationFormat>On-screen Show (16:9)</PresentationFormat>'
    print -r -- "  <Slides>${slide_count}</Slides>"
    print -r -- '  <Notes>0</Notes>'
    print -r -- '  <HiddenSlides>0</HiddenSlides>'
    print -r -- '  <MMClips>0</MMClips>'
    print -r -- '  <ScaleCrop>false</ScaleCrop>'
    print -r -- '  <HeadingPairs>'
    print -r -- '    <vt:vector size="2" baseType="variant">'
    print -r -- '      <vt:variant><vt:lpstr>Slides</vt:lpstr></vt:variant>'
    print -r -- "      <vt:variant><vt:i4>${slide_count}</vt:i4></vt:variant>"
    print -r -- '    </vt:vector>'
    print -r -- '  </HeadingPairs>'
    print -r -- "  <TitlesOfParts><vt:vector size=\"${slide_count}\" baseType=\"lpstr\">"
    local i
    for (( i = 1; i <= slide_count; i++ )); do
      print -r -- "    <vt:lpstr>Slide ${i}</vt:lpstr>"
    done
    print -r -- '  </vt:vector></TitlesOfParts>'
    print -r -- '  <Company/>'
    print -r -- '  <LinksUpToDate>false</LinksUpToDate>'
    print -r -- '  <SharedDoc>false</SharedDoc>'
    print -r -- '  <HyperlinksChanged>false</HyperlinksChanged>'
    print -r -- '  <AppVersion>16.0000</AppVersion>'
    print -r -- '</Properties>'
  } > "$work/docProps/app.xml"

  cat > "$work/docProps/core.xml" <<XML
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>${title}</dc:title>
  <dc:creator>图片转PPTX.command</dc:creator>
  <cp:lastModifiedBy>图片转PPTX.command</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
</cp:coreProperties>
XML

  {
    print -r -- '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    print -r -- '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    print -r -- '                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    print -r -- '                xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
    print -r -- '  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
    print -r -- '  <p:sldIdLst>'
    local i
    for (( i = 1; i <= slide_count; i++ )); do
      print -r -- "    <p:sldId id=\"$((255 + i))\" r:id=\"rId$((i + 1))\"/>"
    done
    print -r -- '  </p:sldIdLst>'
    print -r -- "  <p:sldSz cx=\"${SLIDE_WIDTH_EMU}\" cy=\"${SLIDE_HEIGHT_EMU}\" type=\"wide\"/>"
    print -r -- '  <p:notesSz cx="6858000" cy="9144000"/>'
    print -r -- '  <p:defaultTextStyle/>'
    print -r -- '</p:presentation>'
  } > "$work/ppt/presentation.xml"

  {
    print -r -- '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    print -r -- '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    print -r -- '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>'
    local i
    for (( i = 1; i <= slide_count; i++ )); do
      print -r -- "  <Relationship Id=\"rId$((i + 1))\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide${i}.xml\"/>"
    done
    print -r -- "  <Relationship Id=\"rId$((slide_count + 2))\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps\" Target=\"presProps.xml\"/>"
    print -r -- "  <Relationship Id=\"rId$((slide_count + 3))\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps\" Target=\"viewProps.xml\"/>"
    print -r -- "  <Relationship Id=\"rId$((slide_count + 4))\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles\" Target=\"tableStyles.xml\"/>"
    print -r -- '</Relationships>'
  } > "$work/ppt/_rels/presentation.xml.rels"

  cat > "$work/ppt/presProps.xml" <<'XML'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>
XML

  cat > "$work/ppt/viewProps.xml" <<'XML'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr>
  <p:slideViewPr><p:cSldViewPr><p:cViewPr varScale="1"><p:scale><a:sx n="100" d="100"/><a:sy n="100" d="100"/></p:scale><p:origin x="0" y="0"/></p:cViewPr><p:guideLst/></p:cSldViewPr></p:slideViewPr>
  <p:notesTextViewPr><p:cViewPr><p:scale><a:sx n="100" d="100"/><a:sy n="100" d="100"/></p:scale><p:origin x="0" y="0"/></p:cViewPr></p:notesTextViewPr>
  <p:gridSpacing cx="72008" cy="72008"/>
</p:viewPr>
XML

  cat > "$work/ppt/tableStyles.xml" <<'XML'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>
XML
}

write_theme_and_layouts() {
  local work="$1"

  cat > "$work/ppt/slideMasters/slideMaster1.xml" <<'XML'
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
XML

  cat > "$work/ppt/slideMasters/_rels/slideMaster1.xml.rels" <<'XML'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>
XML

  cat > "$work/ppt/slideLayouts/slideLayout1.xml" <<'XML'
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
XML

  cat > "$work/ppt/slideLayouts/_rels/slideLayout1.xml.rels" <<'XML'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>
XML

  cat > "$work/ppt/theme/theme1.xml" <<'XML'
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
XML
}

write_slide() {
  local work="$1"
  local slide_index="$2"
  local package_name="$3"
  local pic_x="$4"
  local pic_y="$5"
  local pic_cx="$6"
  local pic_cy="$7"

  cat > "$work/ppt/slides/slide${slide_index}.xml" <<XML
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
          <p:cNvPr id="2" name="Picture ${slide_index}"/>
          <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
          <p:nvPr/>
        </p:nvPicPr>
        <p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
        <p:spPr>
          <a:xfrm><a:off x="${pic_x}" y="${pic_y}"/><a:ext cx="${pic_cx}" cy="${pic_cy}"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
      </p:pic>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>
XML

  cat > "$work/ppt/slides/_rels/slide${slide_index}.xml.rels" <<XML
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${package_name}"/>
</Relationships>
XML
}

exif_orientation() {
  # 读取 JPEG 的 EXIF 方向标记(1-8)，读不到、出错或缺少 perl 时返回 1。
  local image_path="$1" result
  [[ -x "$PERL" ]] || { print -r -- 1; return; }
  result="$("$PERL" - "$image_path" 2>/dev/null <<'PERL'
use strict; use warnings;
my $path = $ARGV[0] // "";
open(my $fh, "<:raw", $path) or do { print "1\n"; exit };
my $buf;
unless (read($fh, $buf, 2) == 2 and $buf eq "\xFF\xD8") { print "1\n"; exit }
my $orientation = 1;
my %sof = map { $_ => 1 } (0xC0,0xC1,0xC2,0xC3,0xC5,0xC6,0xC7,0xC9,0xCA,0xCB,0xCD,0xCE,0xCF);
while (1) {
    my $b;
    read($fh, $b, 1) == 1 or last;
    next if $b ne "\xFF";
    do { read($fh, $b, 1) == 1 or last; } while ($b eq "\xFF");
    my $m = ord($b);
    last if $m == 0xD9;
    next if $m == 0x01 or ($m >= 0xD0 and $m <= 0xD7);
    my $lb; read($fh, $lb, 2) == 2 or last;
    my $len = unpack("n", $lb);
    last if $len < 2;
    my $payload = "";
    if ($len > 2) { read($fh, $payload, $len - 2) == $len - 2 or last; }
    if ($m == 0xE1) {
        next if substr($payload, 0, 6) ne "Exif\x00\x00";
        my $tiff = substr($payload, 6);
        last if length($tiff) < 8;
        my $bo = substr($tiff, 0, 2);
        my ($S, $L);
        if    ($bo eq "II") { ($S,$L)=("v","V"); }
        elsif ($bo eq "MM") { ($S,$L)=("n","N"); }
        else { last; }
        last if unpack($S, substr($tiff,2,2)) != 0x2A;
        my $ifd = unpack($L, substr($tiff,4,4));
        last if $ifd + 2 > length($tiff);
        my $cnt = unpack($S, substr($tiff,$ifd,2));
        my $e = $ifd + 2;
        for (my $i=0; $i<$cnt; $i++) {
            last if $e + 12 > length($tiff);
            if (unpack($S, substr($tiff,$e,2)) == 0x0112) {
                my $v = unpack($S, substr($tiff,$e+8,2));
                $orientation = ($v>=1 and $v<=8) ? $v : 1;
                last;
            }
            $e += 12;
        }
        last;
    }
    last if $sof{$m};
}
print "$orientation\n";
PERL
)"
  case "$result" in
    [1-8]) print -r -- "$result" ;;
    *) print -r -- 1 ;;
  esac
}

image_placement() {
  local image_path="$1"
  local info width height pic_x pic_y pic_cx pic_cy

  [[ -x "$SIPS" ]] || fail "这台 Mac 缺少 sips 命令，无法读取图片尺寸。"
  info="$("$SIPS" -g pixelWidth -g pixelHeight "$image_path" 2>/dev/null)" || fail "无法读取图片尺寸：${image_path:t}"
  width="$(print -r -- "$info" | "$AWK" '/pixelWidth:/ {print $2; exit}')"
  height="$(print -r -- "$info" | "$AWK" '/pixelHeight:/ {print $2; exit}')"

  if [[ -z "$width" || -z "$height" || "$width" -le 0 || "$height" -le 0 ]]; then
    fail "图片尺寸异常：${image_path:t}"
  fi

  # sips 报告的是存储像素，不会按 EXIF 方向交换宽高；竖拍照片(方向 5-8)需手动交换。
  if [[ "${image_path:e:l}" == (jpg|jpeg) ]]; then
    local orientation swap
    orientation="$(exif_orientation "$image_path")"
    case "$orientation" in
      5|6|7|8) swap="$width"; width="$height"; height="$swap" ;;
    esac
  fi

  if (( width * SLIDE_HEIGHT_EMU > height * SLIDE_WIDTH_EMU )); then
    pic_cx="$SLIDE_WIDTH_EMU"
    pic_cy=$(( (SLIDE_WIDTH_EMU * height + width / 2) / width ))
    pic_x=0
    pic_y=$(( (SLIDE_HEIGHT_EMU - pic_cy) / 2 ))
  else
    pic_cy="$SLIDE_HEIGHT_EMU"
    pic_cx=$(( (SLIDE_HEIGHT_EMU * width + height / 2) / height ))
    pic_x=$(( (SLIDE_WIDTH_EMU - pic_cx) / 2 ))
    pic_y=0
  fi

  print -r -- "$pic_x $pic_y $pic_cx $pic_cy"
}

scan_images() {
  local scan_dir="$1"
  local image_path ext stem raw_number number_int number

  by_number=()
  numbers=()
  skipped=()

  for image_path in "$scan_dir"/*(.N); do
    ext="${image_path:e:l}"
    case "$ext" in
      jpg|jpeg|png) ;;
      *) continue ;;
    esac

    stem="${image_path:t:r}"
    if [[ "$stem" =~ '(^|[^0-9])([0-9][0-9]+)$' ]]; then
      raw_number="$match[2]"
      number_int=$((10#$raw_number))
      (( number_int < 1 )) && continue
      number="$(printf "%03d" "$number_int")"

      if [[ -n "${by_number[$number]-}" ]]; then
        fail "发现重复编号 ${number}：\n${by_number[$number]:t}\n${image_path:t}\n\n例如 01 和 001 会被视为同一页，请保留其中一个文件后再运行。"
      fi

      by_number[$number]="$image_path"
      numbers+=("$number")
    else
      skipped+=("${image_path:t}")
    fi
  done
}

script_dir="${0:A:h}"
input_dir="${1:-}"
used_script_folder=0
if [[ -z "$input_dir" ]]; then
  input_dir="$script_dir"
  used_script_folder=1
fi

input_dir="${input_dir:a}"
[[ -d "$input_dir" ]] || fail "输入路径不是文件夹：$input_dir"
[[ -x "$ZIP" ]] || fail "这台电脑缺少 zip 命令，无法打包 PPTX。"

typeset -A by_number
typeset -a numbers skipped

scan_images "$input_dir"

if [[ ${#numbers[@]} -eq 0 ]]; then
  if [[ "$used_script_folder" -eq 1 ]]; then
    input_dir="$(choose_input_folder)" || exit 0
    input_dir="${input_dir:a}"
    [[ -d "$input_dir" ]] || fail "输入路径不是文件夹：$input_dir"
    scan_images "$input_dir"
  fi
fi

if [[ ${#numbers[@]} -eq 0 ]]; then
  fail "没有找到文件名以两位及以上数字结尾的 JPG/PNG 图片。"
fi

numbers=("${(@on)numbers}")
slide_count="${#numbers[@]}"
output="$(unique_output_path "$input_dir")"
work="$("$MKTEMP" -d "/tmp/images-to-pptx.XXXXXX")"
trap 'rm -rf "$work"' EXIT

mkdir -p "$work/_rels" \
  "$work/docProps" \
  "$work/ppt/_rels" \
  "$work/ppt/slides/_rels" \
  "$work/ppt/media" \
  "$work/ppt/slideMasters/_rels" \
  "$work/ppt/slideLayouts/_rels" \
  "$work/ppt/theme"

print -r -- "正在生成 PPTX..."
print -r -- "图片文件夹：$input_dir"
print -r -- "输出文件：$output"
print -r -- ""
print -r -- "幻灯片顺序："

write_static_parts "$work" "$slide_count" "${output:t:r}"
write_theme_and_layouts "$work"

slide_index=1
for number in "${numbers[@]}"; do
  source_path="${by_number[$number]}"
  ext="${source_path:e:l}"
  package_name="image${slide_index}.${ext}"
  placement=(${(z)"$(image_placement "$source_path")"})
  print -r -- "  ${(l:3::0:)slide_index}. [$number] ${source_path:t}"
  cp "$source_path" "$work/ppt/media/$package_name" || fail "无法复制图片：${source_path:t}"
  write_slide "$work" "$slide_index" "$package_name" "$placement[1]" "$placement[2]" "$placement[3]" "$placement[4]"
  slide_index=$((slide_index + 1))
done

(cd "$work" && "$ZIP" -qr "$output" .)

print -r -- ""
print -r -- "完成：$output"
print -r -- "共 ${slide_count} 页，图片保持原始比例居中放置。"

"$OPEN" "$output" >/dev/null 2>&1 || true
"$OSASCRIPT" -e "display notification \"已生成 ${slide_count} 页 PPTX，图片未拉伸\" with title \"图片转PPTX完成\"" >/dev/null 2>&1 || true

print -r -- ""
print -r -- "可以关闭这个窗口。"
if [[ -t 0 ]]; then
  read -k 1 -s "?按任意键退出..." || true
fi
