# 更新日志

本项目所有重要变更都记录在本文件中。

## [3.1.0] - 2026-05-30

### 新增

- **三套实现统一改为拖入式终端引导。** Python、macOS zsh、Windows PowerShell 在无参数启动时都会显示标题、步骤、规则和拖入提示；用户把任意一张图片或文件夹拖进终端并回车后，工具会用该路径定位图片文件夹并生成 PPTX。
- **跨平台拖入路径解析。** macOS 支持终端拖入产生的反斜杠转义路径和引号路径；Windows 支持带引号路径；多路径输入时统一取第一个有效路径。

### 变更

- **入口行为调整：脚本不再默认处理脚本所在目录，也不再弹出文件夹选择器。** 现在推荐把脚本放在任意位置，启动后拖入一张图片即可。
- **命令行参数保留为高级用法。** Python 和平台脚本传入图片或文件夹路径时会直接处理并跳过引导；传入图片文件时会自动取其父文件夹。
- **Windows 输出编码改为 UTF-8。** 降低中文提示在终端里乱码的概率。
- **macOS / Windows 取消 GUI 弹窗提示。** 交互集中在终端窗口内，保持两平台一致。

### 文档

- README 改写推荐用法为“启动脚本 → 拖入图片/文件夹 → 回车”。
- README 新增权限提醒：macOS 需要在“系统设置 → 隐私与安全性”里允许打开或授权文件夹访问；Windows 遇到权限不足时可右键以管理员身份运行。

## [3.0.3] - 2026-05-30

### 变更

- **macOS 实现脚本重命名：`图片转PPTX.command` → `图片转PPTX-macOS.command`。** 与 `图片转PPTX-Windows.cmd` 命名对称，文件名直接标明平台。脚本内写入 PPTX 元数据的 `<Application>`／`<dc:creator>`／`<cp:lastModifiedBy>` 三个字段同步更新为新文件名。
- **删除入口脚本 `run_macos.command`。** 它仅是转发到 macOS 实现脚本的 5 行壳脚本，无任何文件/代码引用；macOS 用户现在直接双击 `图片转PPTX-macOS.command` 即可，与 Windows 用法一致。

> 三套实现（Python / macOS zsh / Windows PowerShell）版本号统一升至 3.0.3；本次为文件整理，无功能变化。

## [3.0.2] - 2026-05-29

### 修复

- **macOS：双击运行不再因缺少执行权限而失败。** `run_macos.command` 与 `图片转PPTX.command` 启动时执行 `chmod +x "$0"`（失败静默忽略），首次双击即可运行，无需先在终端手动 `chmod +x`。

> 三套实现（Python / macOS zsh / Windows PowerShell）版本号统一升至 3.0.2。

## [3.0.1] - 2026-05-29

针对 v3.0.0 代码审查（见 [docs/v3-issues.md](docs/v3-issues.md)）发现的问题做修复。三套实现（Python / macOS zsh / Windows PowerShell）行为保持一致。

### 修复

- **macOS（`图片转PPTX.command`）：竖拍照片不再变形。** 此前用 `sips` 读取尺寸，而 `sips` 会忽略 EXIF 方向标记，导致带旋转信息的手机照片在 PPTX 中被压扁或拉伸。现新增 `exif_orientation()`，用系统自带的 `/usr/bin/perl` 解析 EXIF 方向，方向 5–8 时交换宽高；读不到方向或环境无 perl 时安全降级为“不旋转”。输出占位已验证与 Python 版逐像素一致。（审查问题 #1）
- **macOS（`图片转PPTX.command`）：含特殊字符的文件夹名不再生成损坏文件。** `core.xml` 的标题此前直接拼接进 XML，文件夹名含 `& < > " '` 时会产生非法 XML、PowerPoint 无法打开。现新增 `xml_escape()` 做转义。（审查问题 #2）
- **Windows（`图片转PPTX-Windows.cmd`）：演示文稿标题不再硬编码。** `<dc:title>` 此前固定为 “Images to PPTX”，现改用文件夹名，并经 `System.Security.SecurityElement::Escape` 转义，与 Python 版行为一致。（审查问题 #6）

### 变更

- **Python（`images_to_pptx.py`）：递归扫描不再跟随符号链接目录。** `--recursive` 改用 `os.walk(followlinks=False)` 取代 `Path.rglob`，消除不同 Python 版本对符号链接处理的差异，并与非递归模式保持一致。（审查问题 #3）
- **Python：`--strict` 缺号检测改用集合差集**（`sorted(set(range(...)) - numbers)`），更简洁。（审查问题 #4）
- **Python：变量 `ignored_numbered` 重命名为 `skipped_unnumbered`**，名称与“跳过无编号文件”的实际含义相符。（审查问题 #5）
- **Windows：移除冗余的自定义函数 `Join-Path2`**，统一使用 PowerShell 内置 `Join-Path`。（审查问题 #7）

### 说明

- 审查问题 #8（“Python 使用 Tab 缩进”）经复核为**误报**：文件本就使用 4 空格缩进，未做改动。
- 验证：Python 与 macOS zsh 版均已端到端运行（生成 PPTX 并解析校验占位与 XML）；Windows 版因开发机无 PowerShell 未实跑，已通过静态检查（圆括号 / 花括号 / here-string 配平、无 `Join-Path2` 残留）确认。

## [3.0.0] - 2026-05-29

- 初始发布：将按文件名末尾数字（两位及以上）命名的 JPG/PNG 图片合成为 16:9 PPTX，图片保持原始比例居中、不拉伸不裁剪。提供 Python、macOS（zsh）、Windows（PowerShell）三套实现。
