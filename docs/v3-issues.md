# v3.0.0 代码审查记录

> 审查日期：2026-05-29
> 版本标签：`8497de7` — v3.0.0: 初始提交

## ✅ 修复状态（2026-05-29 跟进）

下列问题已在分支 `claude/silly-lehmann-2a55f6` 修复并验证：

| # | 等级 | 问题 | 状态 |
|---|------|------|------|
| 1 | 🔴 | zsh 缺少 EXIF 方向处理 | ✅ 新增 `exif_orientation()`（用 `/usr/bin/perl` 解析 EXIF，缺失时优雅降级为方向 1），方向 5-8 交换宽高；输出占位已验证与 Python 版逐像素一致 |
| 2 | 🔴 | zsh `core.xml` 标题未转义 | ✅ 新增 `xml_escape()`，含 `& < > " '` 的文件夹名现生成合法 XML（已用特殊字符文件夹端到端验证） |
| 3 | 🟡 | Python `rglob` 符号链接不一致 | ✅ 改用 `os.walk(followlinks=False)`，跨版本一致且递归时不跟随符号链接目录 |
| 4 | 🟡 | `ensure_contiguous` 效率 | ✅ 改用集合差集 `sorted(set(range(...)) - numbers)` |
| 5 | 🟢 | 变量命名 `ignored_numbered` | ✅ 重命名为 `skipped_unnumbered` |
| 6 | 🟢 | Windows 标题硬编码 | ✅ 改用文件夹名 + `SecurityElement::Escape` 转义 |
| 7 | 🟢 | PowerShell `Join-Path2` 冗余 | ✅ 移除自定义函数，统一用内置 `Join-Path` |
| 8 | 🟢 | Python Tab 缩进 | ⚠️ **误报**：文件实际已全部使用 4 空格缩进，无需改动 |

> 验证说明：Python 与 macOS zsh 版均已端到端运行（生成 PPTX 并解析校验占位/XML）；Windows 版因本机无 PowerShell 未能运行，已通过静态检查（圆括号/花括号/here-string 配平、无 `Join-Path2` 残留）确认。

## 项目概况

项目包含三个功能等价的实现：

| 文件 | 语言 | 平台 | EXIF 支持 | XML 转义 |
|------|------|------|-----------|----------|
| `images_to_pptx.py` | Python 3 | 跨平台 | ✅ 完整 | ✅ |
| `图片转PPTX.command` | Zsh | macOS | ❌ 缺失 | ❌ 缺标题转义 |
| `图片转PPTX-Windows.cmd` | PowerShell | Windows | ✅ | ✅ |

---

## 🔴 严重（影响功能正确性）

### 1. zsh 脚本：缺少 EXIF 方向处理

**位置**：[`图片转PPTX.command:341-367`](../图片转PPTX.command#L341-L367)

**问题**：`image_placement()` 使用 macOS 的 `sips` 命令读取图片宽高，但 `sips` 不会根据 EXIF orientation 标记交换宽高。

**后果**：竖拍照片（含 EXIF rotation 5-8）在 PPTX 中会被压扁或拉伸变形。

**对比**：
- Python 版 [`images_to_pptx.py:40-67`](../images_to_pptx.py#L40-L67) 完整解析 EXIF APP1 段并处理 orientation 5-8 的宽高交换。
- PowerShell 版 [`图片转PPTX-Windows.cmd:136-142`](../图片转PPTX-Windows.cmd#L136-L142) 通过 `System.Drawing.Image.GetPropertyItem(0x0112)` 同样正确处理。

**修复方向**：zsh 脚本需要引入 EXIF 解析逻辑。可考虑用 `exiftool`（需额外安装）、`mdls -name kMDItemOrientation`（macOS Spotlight 元数据），或直接改为调用 Python 读取图片尺寸。

### 2. zsh 脚本：`core.xml` 标题未做 XML 转义

**位置**：[`图片转PPTX.command:124-137`](../图片转PPTX.command#L124-L137)

**问题**：`${title}` 直接嵌入 XML，如果文件夹名包含 `&`、`<`、`>`、`'`、`"` 等字符，会生成非法 XML。

```bash
# 当前代码
cat > "$work/docProps/core.xml" <<XML
...
  <dc:title>${title}</dc:title>
...
XML
```

**对比**：Python 版 [`images_to_pptx.py:36-37`](../images_to_pptx.py#L36-L37) 使用 `html.escape(value, quote=True)` 正确转义。PowerShell 版不存在此问题（标题硬编码，见下文）。

**后果**：包含特殊字符的文件夹名将导致生成的 PPTX 无法被 PowerPoint 打开。

---

## 🟡 中等（影响兼容性/健壮性）

### 3. Python 版：`rglob` 在不同 Python 版本下行为不一致

**位置**：[`images_to_pptx.py:161`](../images_to_pptx.py#L161)

```python
paths = folder.rglob("*") if recursive else folder.iterdir()
```

**问题**：Python 3.12 起 `Path.rglob()` 默认 `follow_symlinks=True`，而 3.11 及更早版本不跟随符号链接。同时 `iterdir()` 在任何版本都不跟随符号链接目录。

**后果**：
- Python 3.12+ 用户使用 `--recursive` 时可能意外遍历符号链接指向的外部目录
- 非递归模式与递归模式对符号链接的处理不一致

**修复建议**：
```python
paths = folder.rglob("*", follow_symlinks=False) if recursive else folder.iterdir()
```

### 4. Python 版：`ensure_contiguous` 效率问题

**位置**：[`images_to_pptx.py:238-240`](../images_to_pptx.py#L238-L240)

```python
for number in range(lo, hi + 1):
    if number not in numbers:
```

**问题**：当图片编号范围很大但实际文件很少时（例如只有 `IMG_0001.jpg` 和 `IMG_9999.jpg`），循环需要遍历近万次才能收集 30 个缺失样本。时间复杂度为 O(hi-lo)，实际场景不太常见但理论上不够高效。

**修复建议**：
```python
missing_set = set(range(lo, hi + 1)) - numbers
missing_sample = sorted(missing_set)[:30]
```

---

## 🟢 轻微（代码质量/一致性）

### 5. 变量命名误导

**位置**：[`images_to_pptx.py:163`](../images_to_pptx.py#L163)

`ignored_numbered` 字面意思是「被忽略的编号文件」，实际存储的是**没有编号**而被跳过的文件。

**建议**：改为 `skipped_unnumbered`、`ignored_files` 或 `files_without_number`。

### 6. Windows 版标题硬编码

**位置**：[`图片转PPTX-Windows.cmd:254`](../图片转PPTX-Windows.cmd#L254)

```xml
<dc:title>Images to PPTX</dc:title>
```

Python 版使用 `output.stem` 作为标题，Windows 版写死固定文本，行为和另外两个实现不一致。

### 7. PowerShell 脚本路径拼接函数冗余

**位置**：[`图片转PPTX-Windows.cmd:48-51`](../图片转PPTX-Windows.cmd#L48-L51)

脚本中 `Join-Path`（PowerShell 内置）和 `Join-Path2`（自定义 wrapper）功能完全相同但混用，增加不必要的维护成本。

### 8. Python 缩进风格 — ⚠️ 误报（复核后撤销）

原结论称 `images_to_pptx.py` 使用 Tab 缩进。复核发现该文件实际已全部使用 4 空格缩进（`grep -cP '^\t'` 计数为 0），不存在 Tab 缩进问题，无需修改。此条为审查时的误判。

---

## ✅ 已验证正确的部分

- JPEG 解析器正确处理 SOI/EOI、RST marker、多 SOF 类型、多 APP1 段
- PNG 尺寸读取（IHDR chunk 解析）正确
- 编号重复检测（`01` 和 `001` → 同一编号 `1`）逻辑正确
- 图片在 ZIP 中存储为 `ZIP_STORED`（已压缩格式不重复压缩），XML 用 `ZIP_DEFLATED`
- `resolve_output_path` 的文件名冲突降级（`_2`, `_3`...）逻辑正确
- OOXML 结构与 PowerPoint 兼容
- `--strict` 模式编号连续性检查计算正确

---

## 审查总结

| 等级 | 数量 | 涉及文件 |
|------|------|----------|
| 🔴 严重 | 2 | `图片转PPTX.command` |
| 🟡 中等 | 2 | `images_to_pptx.py` |
| 🟢 轻微 | 4 | 全部 |

Python 核心逻辑整体健壮，主要问题集中在 zsh 脚本的 EXIF 和 XML 转义缺失。建议优先修复 zsh 脚本的两个严重问题（尤其是 macOS 是主要目标平台之一的情况下）。
