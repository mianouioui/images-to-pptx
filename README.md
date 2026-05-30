# 图片转 PPTX / Images to PPTX

[中文说明](#中文说明) | [English Documentation](#english-documentation) | [Documentation française](#documentation-française)

---

## 中文说明

### 图片转 PPTX

`images_to_pptx` 用于把一组按数字编号命名的 JPG/PNG 图片快速合成为一个 16:9 PPTX 演示文稿。

图片布局规则：16:9 图片会铺满整页；非 16:9 图片会保持原始比例完整放进页面中央，不拉伸、不裁剪。手机竖拍照片（带 EXIF 旋转信息）会按实际显示方向摆放，不会横躺或变形。

### 图片命名规则

| 规则 | 说明 |
|------|------|
| 文件名末尾 | 必须是**两位及以上数字**（`01`、`001`、`1234` 均可） |
| 排序方式 | 按末尾数字的**数值大小**排序（`2` 排在 `12` 前面） |
| 跳过条件 | 末尾不是数字、或只有一位数字（如 `page1.jpg`）会被跳过 |
| 重复编号 | `01` 和 `001` 视为同一页编号，不能同时存在 |
| 有效格式 | JPG、JPEG、PNG |
| 回退排序 | 如果没有符合编号规则的图片，自动按文件修改时间排序 |

示例文件名：

- `page001.jpg`
- `page02.png`
- `demo_003.jpeg`
- `IMG_1234.jpg`
- `DSC_0042.jpg`

### 运行入口

本项目提供三个运行入口：

- Python 源码入口：`images_to_pptx.py`
- macOS 单文件启动器：`图片转PPTX-macOS.command`
- Windows 启动器：`图片转PPTX-Windows.cmd`

### 运行环境

| 入口 | 运行要求 |
|------|----------|
| `images_to_pptx.py` | Python 3.8+ |
| `图片转PPTX-macOS.command` | macOS + Python 3，双击即可运行 |
| `图片转PPTX-Windows.cmd` | Windows，双击即可运行（内嵌 Python 实现，不依赖外部 `.py` 文件） |

本项目不需要安装第三方 Python 包，也不需要虚拟环境。macOS 和 Windows 入口均已内嵌完整实现，用户无需 Python 知识即可使用。只需在已有 Microsoft 365、PowerPoint、Keynote 或 LibreOffice 的设备上即可打开生成结果。

### macOS 使用方式

`图片转PPTX-macOS.command` 内嵌完整实现。运行时利用系统自带 Python 3 完成图片合成；同时首次启动时会自动为自己赋予执行权限。

1. 双击 `图片转PPTX-macOS.command`
2. 如果 macOS 提示”无法验证开发者”，右键点击文件，选择”打开”，再确认运行
3. 将任意一张图片或文件夹拖入终端窗口，按回车
4. PPTX 会生成在该图片所在的文件夹里

也可以在终端中运行：

```bash
./图片转PPTX-macOS.command /path/to/images
./图片转PPTX-macOS.command /path/to/images/page001.jpg
```

如果 macOS 提示无法访问文件夹，请在”系统设置 → 隐私与安全性 → 文件与文件夹”或”完全磁盘访问权限”中给 Terminal 授权。

### Windows 使用方式

`图片转PPTX-Windows.cmd` 是 Windows 单文件启动入口。双击运行会内嵌执行 Python 实现，不依赖外部文件。

1. 双击 `图片转PPTX-Windows.cmd`
2. 如果遇到 SmartScreen 拦截，选择”更多信息 → 仍要运行”
3. 将任意一张图片或文件夹拖入终端窗口，按回车
4. PPTX 会生成在该图片所在的文件夹里

也可以在命令行中运行：

```cmd
图片转PPTX-Windows.cmd D:\images
图片转PPTX-Windows.cmd D:\images\page001.jpg
```

如果图片位于受保护目录导致权限不足，请右键 `图片转PPTX-Windows.cmd`，选择”以管理员身份运行”，或将图片移到桌面、文档等普通用户目录后再运行。

### Python 源码使用方式

无参数运行会进入拖入式终端引导：

```bash
python3 images_to_pptx.py
```

直接传入图片或图片文件夹路径：

```bash
python3 images_to_pptx.py /path/to/images --open
python3 images_to_pptx.py /path/to/images/page001.jpg --open
```

可用选项：

| 选项 | 说明 |
|------|------|
| 位置参数 | 图片文件或图片文件夹路径；省略时进入拖入式引导 |
| `-o` / `--output` | 指定输出 PPTX 路径 |
| `--title` | 自定义演示文稿标题 |
| `--recursive` | 递归遍历子文件夹 |
| `--dry-run` | 仅检查排序，不生成 PPTX |
| `--open` | 生成后自动用系统默认应用打开 PPTX |
| `--version` | 显示版本号 |

### 输出说明

默认将 PPTX 输出到图片所在文件夹，文件名取自该文件夹名。如果输出文件已存在，会自动生成 `_2`、`_3` 等新文件名。

输出示例：

```text
/path/to/images/
├── page001.jpg
├── page002.jpg
├── page003.png
├── page004.jpg
└── images.pptx        ← 生成的 PPTX
```

### 工作原理

1. 扫描输入文件夹中的 JPG/PNG 图片
2. 优先按文件名末尾数字的数值大小排序；无编号时按文件修改时间排序
3. 检查图片 EXIF 旋转信息，自动修正显示方向
4. 按 16:9 宽屏尺寸（12192000 × 6858000 EMU）创建幻灯片
5. 使用 Open XML 标准内建元素构建 PPTX，**不依赖 python-pptx 等第三方库**
6. 图片以原始格式直接存入 PPTX 内部，不做二次压缩
7. 每张图片根据自身比例居中放置于幻灯片中

### 开源协议

MIT

---

## English Documentation

### Images to PPTX

`images_to_pptx` rapidly assembles a set of sequentially-numbered JPG/PNG images into a single 16:9 PPTX presentation.

Layout rules: 16:9 images fill the entire slide; non-16:9 images are centered at their original aspect ratio — no stretching, no cropping. Phone portrait shots with EXIF rotation metadata are displayed in their correct orientation.

### Naming Rules

| Rule | Detail |
|------|--------|
| Filename suffix | Must end with **two or more digits** (`01`, `001`, `1234` all accepted) |
| Sort order | By **numeric value** of the trailing digits (`2` comes before `12`) |
| Excluded files | Files whose suffix is not a number, or has only one digit (e.g. `page1.jpg`) |
| Duplicate numbers | `01` and `001` are treated as the same page number and cannot coexist |
| Supported formats | JPG, JPEG, PNG |
| Fallback sort | When no images match the numbering rule, sort by file modification time automatically |

### Runtime Entry Points

This project provides three runtime entry points:

- Python source entry point: `images_to_pptx.py`
- macOS single-file launcher: `图片转PPTX-macOS.command`
- Windows launcher: `图片转PPTX-Windows.cmd`

### Requirements

| Entry Point | Requirement |
|-------------|-------------|
| `images_to_pptx.py` | Python 3.8+ |
| `图片转PPTX-macOS.command` | macOS + Python 3; double-click to run |
| `图片转PPTX-Windows.cmd` | Windows; double-click to run (embeds Python implementation; no external `.py` file needed) |

No third-party Python packages or virtual environment are required. Both the macOS and Windows entry points embed the full implementation — no Python knowledge is needed. Opening the generated PPTX only requires Microsoft 365, PowerPoint, Keynote, or LibreOffice.

### macOS Usage

`图片转PPTX-macOS.command` embeds the full implementation, running the system Python 3 for image assembly. On first launch it automatically grants itself execute permission.

1. Double-click `图片转PPTX-macOS.command`
2. If macOS shows an unidentified-developer warning, right-click the file, choose Open, and confirm
3. Drag any image or folder into the terminal window and press Enter
4. The PPTX is generated in the same folder as the source image

Terminal usage:

```bash
./图片转PPTX-macOS.command /path/to/images
./图片转PPTX-macOS.command /path/to/images/page001.jpg
```

If macOS denies folder access, grant Terminal permission under “System Settings → Privacy & Security → Files and Folders” or “Full Disk Access”.

### Windows Usage

`图片转PPTX-Windows.cmd` is the Windows single-file entry point. Double-clicking runs the embedded Python implementation with no external dependencies.

1. Double-click `图片转PPTX-Windows.cmd`
2. If SmartScreen blocks execution, choose “More info → Run anyway”
3. Drag any image or folder into the terminal window and press Enter
4. The PPTX is generated in the same folder as the source image

Command-line usage:

```cmd
图片转PPTX-Windows.cmd D:\images
图片转PPTX-Windows.cmd D:\images\page001.jpg
```

If images are in a protected directory, right-click `图片转PPTX-Windows.cmd` and choose “Run as administrator”, or move the images to a regular user directory such as Desktop or Documents.

### Python Source Usage

Running without arguments enters the interactive drag-and-drop terminal mode:

```bash
python3 images_to_pptx.py
```

Pass an image or folder path directly:

```bash
python3 images_to_pptx.py /path/to/images --open
python3 images_to_pptx.py /path/to/images/page001.jpg --open
```

Available options:

| Option | Description |
|--------|-------------|
| Positional arg | Image file or folder path; enters interactive mode when omitted |
| `-o` / `--output` | Specify output PPTX path |
| `--title` | Custom presentation title |
| `--recursive` | Recurse into subdirectories |
| `--dry-run` | Check sort order only, do not generate PPTX |
| `--open` | Automatically open the PPTX with the system default application |
| `--version` | Display version number |

### Output

By default the PPTX is saved in the source image folder, named after that folder. If the output file already exists, `_2`, `_3` suffixes are appended automatically.

Output example:

```text
/path/to/images/
├── page001.jpg
├── page002.jpg
├── page003.png
├── page004.jpg
└── images.pptx        ← generated PPTX
```

### How It Works

1. Scan the input folder for JPG/PNG images
2. Sort by trailing-digit numeric value when available; fall back to file modification time otherwise
3. Read EXIF rotation metadata and auto-correct display orientation
4. Create slides at 16:9 widescreen dimensions (12192000 × 6858000 EMU)
5. Build the PPTX using Open XML standard elements — **no third-party libraries such as python-pptx required**
6. Store images in their original format inside the PPTX package with no re-compression
7. Center each image on its slide according to its aspect ratio

### License

MIT

---

## Documentation française

### Images to PPTX

`images_to_pptx` assemble rapidement un ensemble d'images JPG/PNG numérotées en une seule présentation PPTX 16:9.

Règles de mise en page : les images 16:9 remplissent toute la diapositive ; les images non 16:9 sont centrées en conservant leur ratio d'origine — aucun étirement, aucun recadrage. Les photos de téléphone en mode portrait avec métadonnées de rotation EXIF sont affichées dans leur orientation correcte.

### Règles de nommage

| Règle | Détail |
|-------|--------|
| Suffixe du nom | Doit se terminer par **deux chiffres ou plus** (`01`, `001`, `1234` sont acceptés) |
| Ordre de tri | Par **valeur numérique** des derniers chiffres (`2` vient avant `12`) |
| Fichiers exclus | Fichiers dont le suffixe n'est pas un nombre, ou n'a qu'un seul chiffre (ex. `page1.jpg`) |
| Numéros en double | `01` et `001` sont traités comme le même numéro de page et ne peuvent coexister |
| Formats supportés | JPG, JPEG, PNG |
| Tri de secours | Quand aucune image ne respecte la règle de nommage, tri par date de modification du fichier |

### Points d'entrée

Le projet fournit trois points d'entrée :

- Script Python source : `images_to_pptx.py`
- Lanceur macOS en fichier unique : `图片转PPTX-macOS.command`
- Lanceur Windows : `图片转PPTX-Windows.cmd`

### Prérequis

| Point d'entrée | Prérequis |
|----------------|-----------|
| `images_to_pptx.py` | Python 3.8+ |
| `图片转PPTX-macOS.command` | macOS + Python 3 ; double-cliquez pour lancer |
| `图片转PPTX-Windows.cmd` | Windows ; double-cliquez pour lancer (intègre l'implémentation Python ; aucun fichier `.py` externe requis) |

Aucun paquet Python tiers ni environnement virtuel n'est nécessaire. Les points d'entrée macOS et Windows intègrent l'implémentation complète — aucune connaissance de Python n'est requise. L'ouverture du PPTX généré nécessite uniquement Microsoft 365, PowerPoint, Keynote ou LibreOffice.

### Utilisation sur macOS

`图片转PPTX-macOS.command` intègre l'implémentation complète et utilise le Python 3 du système pour l'assemblage des images. Au premier lancement, il s'accorde automatiquement la permission d'exécution.

1. Double-cliquez sur `图片转PPTX-macOS.command`
2. Si macOS affiche un avertissement de développeur non identifié, faites un clic droit sur le fichier, choisissez Ouvrir, puis confirmez
3. Glissez n'importe quelle image ou dossier dans la fenêtre du terminal et appuyez sur Entrée
4. Le PPTX est généré dans le même dossier que l'image source

Utilisation dans le terminal :

```bash
./图片转PPTX-macOS.command /path/to/images
./图片转PPTX-macOS.command /path/to/images/page001.jpg
```

Si macOS refuse l'accès au dossier, accordez l'autorisation à Terminal dans « Réglages Système → Confidentialité et sécurité → Fichiers et dossiers » ou « Accès complet au disque ».

### Utilisation sur Windows

`图片转PPTX-Windows.cmd` est le point d'entrée Windows en fichier unique. Un double-clic exécute l'implémentation Python intégrée sans dépendance externe.

1. Double-cliquez sur `图片转PPTX-Windows.cmd`
2. Si SmartScreen bloque l'exécution, choisissez « Plus d'infos → Exécuter quand même »
3. Glissez n'importe quelle image ou dossier dans la fenêtre du terminal et appuyez sur Entrée
4. Le PPTX est généré dans le même dossier que l'image source

Utilisation en ligne de commande :

```cmd
图片转PPTX-Windows.cmd D:\images
图片转PPTX-Windows.cmd D:\images\page001.jpg
```

Si les images se trouvent dans un répertoire protégé, faites un clic droit sur `图片转PPTX-Windows.cmd` et choisissez « Exécuter en tant qu'administrateur », ou déplacez les images vers un dossier utilisateur standard tel que le Bureau ou Documents.

### Utilisation du script Python

Lancer sans argument ouvre le mode interactif par glisser-déposer dans le terminal :

```bash
python3 images_to_pptx.py
```

Passez directement un chemin d'image ou de dossier :

```bash
python3 images_to_pptx.py /path/to/images --open
python3 images_to_pptx.py /path/to/images/page001.jpg --open
```

Options disponibles :

| Option | Description |
|--------|-------------|
| Argument positionnel | Chemin du fichier image ou du dossier ; entre en mode interactif si omis |
| `-o` / `--output` | Spécifier le chemin du PPTX de sortie |
| `--title` | Titre personnalisé de la présentation |
| `--recursive` | Parcourir les sous-dossiers récursivement |
| `--dry-run` | Vérifier l'ordre de tri uniquement, sans générer le PPTX |
| `--open` | Ouvrir automatiquement le PPTX avec l'application par défaut du système |
| `--version` | Afficher le numéro de version |

### Sortie

Par défaut, le PPTX est enregistré dans le dossier des images sources et nommé d'après ce dossier. Si le fichier de sortie existe déjà, des suffixes `_2`, `_3` sont ajoutés automatiquement.

Exemple de sortie :

```text
/path/to/images/
├── page001.jpg
├── page002.jpg
├── page003.png
├── page004.jpg
└── images.pptx        ← PPTX généré
```

### Fonctionnement

1. Analyser le dossier d'entrée à la recherche d'images JPG/PNG
2. Trier par valeur numérique des derniers chiffres du nom quand disponibles ; sinon, trier par date de modification du fichier
3. Lire les métadonnées de rotation EXIF et corriger automatiquement l'orientation d'affichage
4. Créer des diapositives au format écran large 16:9 (12192000 × 6858000 EMU)
5. Construire le PPTX en utilisant les éléments standard Open XML — **aucune bibliothèque tierce telle que python-pptx n'est requise**
6. Stocker les images dans leur format d'origine à l'intérieur du paquet PPTX sans recompression
7. Centrer chaque image sur sa diapositive en fonction de son ratio d'aspect

### Licence

MIT
