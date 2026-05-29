#!/usr/bin/env python3
from __future__ import annotations

import argparse
import html
import os
import re
import sys
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZIP_STORED, ZipFile

__version__ = "3.0.2"

SLIDE_WIDTH_EMU = 12192000
SLIDE_HEIGHT_EMU = 6858000
SUPPORTED_SUFFIXES = {
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png": "image/png",
}
TRAILING_NUMBER_RE = re.compile(r"(?:^|\D)(\d{2,})$")


@dataclass(frozen=True)
class ImageItem:
    path: Path
    number: int
    package_name: str
    x: int
    y: int
    cx: int
    cy: int


def xml_escape(value: str) -> str:
    return html.escape(value, quote=True)


def exif_orientation(app1_payload: bytes) -> int:
    """从 JPEG APP1(Exif) 段读取方向标记(1-8)，读不到时返回 1。"""
    if app1_payload[:6] != b"Exif\x00\x00":
        return 1
    tiff = app1_payload[6:]
    if len(tiff) < 8:
        return 1
    if tiff[:2] == b"II":
        byte_order = "little"
    elif tiff[:2] == b"MM":
        byte_order = "big"
    else:
        return 1
    if int.from_bytes(tiff[2:4], byte_order) != 0x2A:
        return 1
    ifd_offset = int.from_bytes(tiff[4:8], byte_order)
    if ifd_offset + 2 > len(tiff):
        return 1
    entry_count = int.from_bytes(tiff[ifd_offset:ifd_offset + 2], byte_order)
    entry = ifd_offset + 2
    for _ in range(entry_count):
        if entry + 12 > len(tiff):
            break
        if int.from_bytes(tiff[entry:entry + 2], byte_order) == 0x0112:
            value = int.from_bytes(tiff[entry + 8:entry + 10], byte_order)
            return value if 1 <= value <= 8 else 1
        entry += 12
    return 1


def read_image_size(path: Path) -> tuple[int, int]:
    suffix = path.suffix.lower()
    with path.open("rb") as image:
        if suffix == ".png":
            header = image.read(24)
            if header[:8] != b"\x89PNG\r\n\x1a\n" or header[12:16] != b"IHDR":
                raise ValueError("不是有效的 PNG 文件")
            return int.from_bytes(header[16:20], "big"), int.from_bytes(header[20:24], "big")

        if suffix in {".jpg", ".jpeg"}:
            if image.read(2) != b"\xff\xd8":
                raise ValueError("不是有效的 JPEG 文件")

            sof_markers = {
                0xC0, 0xC1, 0xC2, 0xC3,
                0xC5, 0xC6, 0xC7,
                0xC9, 0xCA, 0xCB,
                0xCD, 0xCE, 0xCF,
            }

            orientation = 1
            width = height = None

            while True:
                marker_start = image.read(1)
                if not marker_start:
                    break
                if marker_start != b"\xff":
                    continue

                marker = image.read(1)
                while marker == b"\xff":
                    marker = image.read(1)
                if not marker:
                    break

                marker_value = marker[0]
                if marker_value in {0xD8, 0xD9, 0x01} or 0xD0 <= marker_value <= 0xD7:
                    continue

                segment_length_data = image.read(2)
                if len(segment_length_data) != 2:
                    break
                segment_length = int.from_bytes(segment_length_data, "big")
                if segment_length < 2:
                    raise ValueError("JPEG 段长度异常")

                if marker_value == 0xE1:
                    payload = image.read(segment_length - 2)
                    orientation = exif_orientation(payload)
                    continue

                if marker_value in sof_markers:
                    data = image.read(5)
                    if len(data) != 5:
                        break
                    height = int.from_bytes(data[1:3], "big")
                    width = int.from_bytes(data[3:5], "big")
                    break

                image.seek(segment_length - 2, 1)

            if width is None or height is None:
                raise ValueError("无法读取图片尺寸")
            # 方向 5-8 表示拍摄时旋转了 90°，显示尺寸需交换宽高
            if orientation in (5, 6, 7, 8):
                return height, width
            return width, height

    raise ValueError("无法读取图片尺寸")


def image_placement(width: int, height: int) -> tuple[int, int, int, int]:
    if width <= 0 or height <= 0:
        raise ValueError("图片尺寸异常")

    if width * SLIDE_HEIGHT_EMU > height * SLIDE_WIDTH_EMU:
        cx = SLIDE_WIDTH_EMU
        cy = (SLIDE_WIDTH_EMU * height + width // 2) // width
        x = 0
        y = (SLIDE_HEIGHT_EMU - cy) // 2
    else:
        cy = SLIDE_HEIGHT_EMU
        cx = (SLIDE_HEIGHT_EMU * width + height // 2) // height
        x = (SLIDE_WIDTH_EMU - cx) // 2
        y = 0

    return x, y, cx, cy


def iter_candidate_paths(folder: Path, recursive: bool):
    """遍历待处理的文件。递归模式用 os.walk 且不跟随符号链接目录，与非递归的
    iterdir() 行为一致，也避免不同 Python 版本 Path.rglob 对符号链接处理的差异。"""
    if not recursive:
        yield from folder.iterdir()
        return
    for root, _dirs, files in os.walk(folder, followlinks=False):
        root_path = Path(root)
        for name in files:
            yield root_path / name


def find_images(folder: Path, recursive: bool) -> list[ImageItem]:
    paths = iter_candidate_paths(folder, recursive)
    images: list[tuple[int, Path, str]] = []  # (number, path, suffix)
    skipped_unnumbered: list[Path] = []

    for path in paths:
        if not path.is_file():
            continue

        suffix = path.suffix.lower()
        if suffix not in SUPPORTED_SUFFIXES:
            continue

        match = TRAILING_NUMBER_RE.search(path.stem)
        if not match:
            skipped_unnumbered.append(path)
            continue

        number = int(match.group(1))
        if number < 1:
            skipped_unnumbered.append(path)
            continue

        images.append((number, path, suffix))

    images.sort(key=lambda item: (item[0], item[1].name.lower()))

    seen: dict[int, Path] = {}
    duplicates: list[tuple[int, Path, Path]] = []
    for number, path, _suffix in images:
        if number in seen:
            duplicates.append((number, seen[number], path))
        else:
            seen[number] = path

    if duplicates:
        lines = ["发现重复编号，无法判断幻灯片顺序："]
        for number, first, second in duplicates:
            lines.append(f"  {number:03d}: {first.name} / {second.name}")
        raise SystemExit("\n".join(lines))

    result: list[ImageItem] = []
    for index, (number, path, suffix) in enumerate(images, start=1):
        try:
            width, height = read_image_size(path)
            x, y, cx, cy = image_placement(width, height)
        except ValueError as exc:
            raise SystemExit(f"无法读取图片尺寸：{path.name}\n{exc}") from exc

        result.append(
            ImageItem(
                path=path,
                number=number,
                package_name=f"image{index}{suffix}",
                x=x,
                y=y,
                cx=cx,
                cy=cy,
            )
        )

    if skipped_unnumbered:
        print("已跳过文件名末尾不是数字（至少两位）的图片：", file=sys.stderr)
        for path in skipped_unnumbered[:20]:
            print(f"  {path.name}", file=sys.stderr)
        if len(skipped_unnumbered) > 20:
            print(f"  ... 还有 {len(skipped_unnumbered) - 20} 个", file=sys.stderr)

    return result


def ensure_contiguous(images: list[ImageItem]) -> None:
    numbers = {image.number for image in images}
    lo, hi = min(numbers), max(numbers)
    total_missing = (hi - lo + 1) - len(numbers)
    if total_missing <= 0:
        return
    missing = sorted(set(range(lo, hi + 1)) - numbers)
    sample = ", ".join(f"{number:03d}" for number in missing[:30])
    suffix = "" if total_missing <= 30 else f" ... 还有 {total_missing - 30} 个"
    raise SystemExit(f"strict 模式发现缺失编号：{sample}{suffix}")


def resolve_output_path(input_folder: Path, requested: Path | None, force: bool) -> Path:
    if requested is None:
        base_name = input_folder.name or "images"
        output = input_folder / f"{base_name}.pptx"
    else:
        output = requested.expanduser()
        if output.exists() and output.is_dir():
            output = output / f"{input_folder.name or 'images'}.pptx"
        elif output.suffix.lower() != ".pptx":
            output = output.with_suffix(".pptx")

    output = output.resolve()
    output.parent.mkdir(parents=True, exist_ok=True)

    if force or not output.exists():
        return output

    stem = output.stem
    suffix = output.suffix
    parent = output.parent
    for index in range(2, 1000):
        candidate = parent / f"{stem}_{index}{suffix}"
        if not candidate.exists():
            return candidate

    raise SystemExit(f"输出文件已存在且无法生成新文件名：{output}")


def content_types_xml(slide_count: int) -> str:
    slide_overrides = "\n".join(
        f'  <Override PartName="/ppt/slides/slide{i}.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for i in range(1, slide_count + 1)
    )
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
{slide_overrides}
</Types>
'''


def root_rels_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
'''


def app_xml(slide_count: int) -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>images_to_pptx.py</Application>
  <PresentationFormat>On-screen Show (16:9)</PresentationFormat>
  <Slides>{slide_count}</Slides>
  <Notes>0</Notes>
  <HiddenSlides>0</HiddenSlides>
  <MMClips>0</MMClips>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs>
    <vt:vector size="2" baseType="variant">
      <vt:variant><vt:lpstr>Slides</vt:lpstr></vt:variant>
      <vt:variant><vt:i4>{slide_count}</vt:i4></vt:variant>
    </vt:vector>
  </HeadingPairs>
  <TitlesOfParts>
    <vt:vector size="{slide_count}" baseType="lpstr">
      {''.join(f'<vt:lpstr>Slide {i}</vt:lpstr>' for i in range(1, slide_count + 1))}
    </vt:vector>
  </TitlesOfParts>
  <Company/>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0000</AppVersion>
</Properties>
'''


def core_xml(title: str) -> str:
    now = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    safe_title = xml_escape(title)
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>{safe_title}</dc:title>
  <dc:creator>images_to_pptx.py</dc:creator>
  <cp:lastModifiedBy>images_to_pptx.py</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
</cp:coreProperties>
'''


def presentation_xml(slide_count: int) -> str:
    slide_ids = "\n".join(
        f'    <p:sldId id="{255 + i}" r:id="rId{i + 1}"/>' for i in range(1, slide_count + 1)
    )
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
{slide_ids}
  </p:sldIdLst>
  <p:sldSz cx="{SLIDE_WIDTH_EMU}" cy="{SLIDE_HEIGHT_EMU}" type="wide"/>
  <p:notesSz cx="6858000" cy="9144000"/>
  <p:defaultTextStyle/>
</p:presentation>
'''


def presentation_rels_xml(slide_count: int) -> str:
    slide_rels = "\n".join(
        f'  <Relationship Id="rId{i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i}.xml"/>'
        for i in range(1, slide_count + 1)
    )
    pres_props_id = slide_count + 2
    view_props_id = slide_count + 3
    table_styles_id = slide_count + 4
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
{slide_rels}
  <Relationship Id="rId{pres_props_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>
  <Relationship Id="rId{view_props_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>
  <Relationship Id="rId{table_styles_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>
</Relationships>
'''


def group_shape_xml() -> str:
    return '''      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>'''


def slide_xml(slide_index: int, image: ImageItem) -> str:
    name = xml_escape(image.path.name)
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
{group_shape_xml()}
      <p:pic>
        <p:nvPicPr>
          <p:cNvPr id="2" name="Picture {slide_index}" descr="{name}"/>
          <p:cNvPicPr>
            <a:picLocks noChangeAspect="1"/>
          </p:cNvPicPr>
          <p:nvPr/>
        </p:nvPicPr>
        <p:blipFill>
          <a:blip r:embed="rId2"/>
          <a:stretch>
            <a:fillRect/>
          </a:stretch>
        </p:blipFill>
        <p:spPr>
          <a:xfrm>
            <a:off x="{image.x}" y="{image.y}"/>
            <a:ext cx="{image.cx}" cy="{image.cy}"/>
          </a:xfrm>
          <a:prstGeom prst="rect">
            <a:avLst/>
          </a:prstGeom>
        </p:spPr>
      </p:pic>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sld>
'''


def slide_rels_xml(image: ImageItem) -> str:
    target = f"../media/{image.package_name}"
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{target}"/>
</Relationships>
'''


def slide_master_xml() -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg>
      <p:bgRef idx="1001">
        <a:schemeClr val="bg1"/>
      </p:bgRef>
    </p:bg>
    <p:spTree>
{group_shape_xml()}
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId1"/>
  </p:sldLayoutIdLst>
  <p:txStyles>
    <p:titleStyle/>
    <p:bodyStyle/>
    <p:otherStyle/>
  </p:txStyles>
</p:sldMaster>
'''


def slide_master_rels_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>
'''


def slide_layout_xml() -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             type="blank" preserve="1">
  <p:cSld name="Blank">
    <p:spTree>
{group_shape_xml()}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sldLayout>
'''


def slide_layout_rels_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>
'''


def theme_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
      <a:majorFont>
        <a:latin typeface="Aptos Display"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Aptos"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs>
            <a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:lin ang="5400000" scaled="0"/>
        </a:gradFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs>
            <a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:lin ang="5400000" scaled="0"/>
        </a:gradFill>
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
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs>
            <a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:lin ang="5400000" scaled="0"/>
        </a:gradFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>
'''


def pres_props_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>
'''


def view_props_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:normalViewPr>
    <p:restoredLeft sz="15620"/>
    <p:restoredTop sz="94660"/>
  </p:normalViewPr>
  <p:slideViewPr>
    <p:cSldViewPr>
      <p:cViewPr varScale="1">
        <p:scale>
          <a:sx n="100" d="100"/>
          <a:sy n="100" d="100"/>
        </p:scale>
        <p:origin x="0" y="0"/>
      </p:cViewPr>
      <p:guideLst/>
    </p:cSldViewPr>
  </p:slideViewPr>
  <p:notesTextViewPr>
    <p:cViewPr>
      <p:scale>
        <a:sx n="100" d="100"/>
        <a:sy n="100" d="100"/>
      </p:scale>
      <p:origin x="0" y="0"/>
    </p:cViewPr>
  </p:notesTextViewPr>
  <p:gridSpacing cx="72008" cy="72008"/>
</p:viewPr>
'''


def table_styles_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>
'''


def build_pptx(images: list[ImageItem], output: Path, title: str) -> None:
    with ZipFile(output, "w", ZIP_DEFLATED) as pptx:
        pptx.writestr("[Content_Types].xml", content_types_xml(len(images)))
        pptx.writestr("_rels/.rels", root_rels_xml())
        pptx.writestr("docProps/app.xml", app_xml(len(images)))
        pptx.writestr("docProps/core.xml", core_xml(title))
        pptx.writestr("ppt/presentation.xml", presentation_xml(len(images)))
        pptx.writestr("ppt/_rels/presentation.xml.rels", presentation_rels_xml(len(images)))
        pptx.writestr("ppt/presProps.xml", pres_props_xml())
        pptx.writestr("ppt/viewProps.xml", view_props_xml())
        pptx.writestr("ppt/tableStyles.xml", table_styles_xml())
        pptx.writestr("ppt/slideMasters/slideMaster1.xml", slide_master_xml())
        pptx.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", slide_master_rels_xml())
        pptx.writestr("ppt/slideLayouts/slideLayout1.xml", slide_layout_xml())
        pptx.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", slide_layout_rels_xml())
        pptx.writestr("ppt/theme/theme1.xml", theme_xml())

        for index, image in enumerate(images, start=1):
            pptx.writestr(f"ppt/slides/slide{index}.xml", slide_xml(index, image))
            pptx.writestr(f"ppt/slides/_rels/slide{index}.xml.rels", slide_rels_xml(image))
            pptx.write(image.path, f"ppt/media/{image.package_name}", compress_type=ZIP_STORED)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="把按文件名末尾数字（两位及以上，如 01、001、1234）命名的 JPG/PNG 图片快速合成为 16:9 PPTX。"
    )
    parser.add_argument(
        "input_folder",
        nargs="?",
        default=".",
        type=Path,
        help="图片所在文件夹，默认是当前文件夹。",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="输出 PPTX 路径。默认输出到图片文件夹内，文件名为“文件夹名.pptx”。",
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="递归扫描子文件夹。",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="要求编号从第一张到最后一张连续，中间不能缺号。",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="只预览排序结果，不生成 PPTX。",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="如果输出文件已存在，直接覆盖。",
    )
    parser.add_argument(
        "--open",
        action="store_true",
        help="生成后用 macOS 默认应用打开 PPTX。",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"images_to_pptx.py v{__version__}",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    input_folder = args.input_folder.expanduser().resolve()

    if not input_folder.exists():
        print(f"图片文件夹不存在：{input_folder}", file=sys.stderr)
        return 2
    if not input_folder.is_dir():
        print(f"输入路径不是文件夹：{input_folder}", file=sys.stderr)
        return 2

    images = find_images(input_folder, args.recursive)
    if not images:
        print("没有找到文件名以两位及以上数字结尾的 JPG/PNG 图片。", file=sys.stderr)
        return 1

    if args.strict:
        ensure_contiguous(images)

    print("幻灯片顺序：")
    for index, image in enumerate(images, start=1):
        print(f"  {index:03d}. [{image.number:03d}] {image.path.name}")

    if args.dry_run:
        print(f"\n共 {len(images)} 张图片；dry-run 模式未生成文件。")
        return 0

    output = resolve_output_path(input_folder, args.output, args.force)
    build_pptx(images, output, title=output.stem)
    print(f"\n已生成：{output}")
    print(f"共 {len(images)} 页，16:9 宽屏，图片保持原始比例居中放置。")

    if args.open:
        import subprocess

        subprocess.run(["open", str(output)], check=False)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
