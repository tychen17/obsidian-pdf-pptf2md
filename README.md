# Obsidian PDF / PPT 转 Markdown 插件（pdf2md）

右键一键把知识库里的 **PDF / PPT** 转换为 Markdown：图片自动存入源文件同目录的 `attachments/`，Markdown 内用相对路径引用，Obsidian 中直接可见。

适合大量「PPT 转出的 PDF」场景：每页可渲染为一张整页图片，正文文字按阅读顺序提取。

## 功能特性

- **右键 PDF →「转换为 Markdown」**：提取文字 + 嵌入图片（自动去重、忽略过小图片）
- **右键 PDF →「转换为 Markdown（每页渲染为图片）」**：每页再渲染一张整页 PNG 放在页首，适合 PPT 转出的 PDF
- **右键 .ppt / .pptx →「转换为 Markdown」**：每页幻灯片渲染为整页 PNG + 文字、表格（转成 Markdown 表格）、嵌入图片提取
- 图片统一存入源文件同目录 `attachments/`，md 内用相对路径 `![](attachments/xxx.png)` 引用
- **增量转换**：已转换且未变更的文件自动跳过，重复运行只会处理新增/更新的文件
- **保护模式**：同名且非本插件生成的 md 不会被覆盖（除非使用 `--force`）
- 命令面板支持（Ctrl+P）
- 依赖缺失时自动安装（pymupdf / python-pptx / pywin32）
- 可设置：Python 路径、OCR 开关与语言、忽略图片最小尺寸

## 目录结构

```
obsidian-pdf2md/
├── manifest.json          # 插件清单（Obsidian 插件元信息）
├── main.js                # 插件主程序（右键菜单 / 命令面板 / 调用脚本）
├── obsidian_pdf2md.py     # PDF → Markdown 转换脚本
├── obsidian_ppt2md.py     # PPT → Markdown 转换脚本
├── data.example.json      # 可选配置模板（需要时复制为 data.json 使用）
└── README.md
```

## 安装

1. **准备 Python**（可选但推荐，详见「Python 环境」一节）：安装 Python 3.9+。
2. **放入插件目录**：将本仓库整个文件夹放到知识库的插件目录下：

   ```
   你的知识库/.obsidian/plugins/pdf2md/
   ```

   > 插件文件夹名需要是 `pdf2md`（与 `manifest.json` 中的 `id` 一致）。
3. **重启 Obsidian**（或按 `Ctrl+R` 重新加载），进入「设置 → 第三方插件」启用 **pdf2md**。

## 使用

- **文件树或标签页中右键** PDF / PPT → 选择对应转换菜单项；
- **命令面板**（Ctrl+P）→「将当前打开的 PDF / PPT 转换为 Markdown」；
- 转换完成后，源文件同目录出现同名 `.md` 与 `attachments/` 图片文件夹。

## Python 环境

插件通过命令行调用 Python 执行转换脚本，解释器按以下顺序选择：

1. **设置 → PDF/PPT 转 Markdown →「Python 路径」**中填写的 `python.exe` 完整路径；
2. 插件目录下自带的 `.venv`（可选，见下节）；
3. 系统 PATH 中的 `python`。

依赖缺失时插件会自动执行 `pip install` 并重试一次，无需手动安装。

常见问题：

| 现象 | 处理 |
| --- | --- |
| 提示「未找到 Python」 | 安装 Python 并勾选加入 PATH，或在设置里填写 python.exe 完整路径 |
| 提示「禁止安装包（externally managed）」 | 当前解释器不允许装包（如 uv 托管解释器），请换一个 Python，或按下面方法给插件建 .venv |
| 提示「插件脚本缺失」 | 插件文件不完整，请重新完整复制本仓库到插件目录 |

### 可选：给插件内置 Python 环境（.venv）

不依赖系统 Python 的推荐做法——在插件目录创建专属虚拟环境：

```bat
python -m venv "你的知识库\.obsidian\plugins\pdf2md\.venv"
"你的知识库\.obsidian\plugins\pdf2md\.venv\Scripts\python.exe" -m pip install pymupdf python-pptx pywin32
```

创建后插件会自动优先使用该环境。

## PPT 渲染说明

- 每页幻灯片渲染为整页 PNG（`attachments/xxx-slide-01.png`…），放在每页开头；
- 渲染引擎优先级：**Microsoft PowerPoint（COM）** → **LibreOffice（soffice）** → 不渲染；
- 有 Office 或 LibreOffice 时：完整渲染；都没有时自动退回「文字 + 嵌入图片」模式；
- 旧版 `.ppt` 需要 Office（python-pptx 不支持旧格式），建议另存为 `.pptx`。

## 命令行批量转换（不装插件也能用）

```bat
:: PDF
python obsidian_pdf2md.py                          :: 转换当前目录下的全部 PDF
python obsidian_pdf2md.py "D:\notes" --force      :: 转换指定目录并强制重新转换
python obsidian_pdf2md.py --dry-run               :: 仅预览将处理的文件，不写入
python obsidian_pdf2md.py --ocr                   :: 对扫描版 PDF 启用 OCR

:: PPT
python obsidian_ppt2md.py                          :: 转换当前目录下的全部 PPT
python obsidian_ppt2md.py "D:\notes\a.pptx"        :: 转换单个文件
```

### obsidian_pdf2md.py 参数

| 参数 | 说明 |
| --- | --- |
| `target` | 目标目录或单个 PDF（默认当前目录） |
| `--force` | 强制重新转换，允许覆盖同名 md |
| `--workers N` | 并行线程数（默认 4） |
| `--min-size N` | 忽略宽或高小于 N pt 的图片（默认 40） |
| `--ocr` | 对无文本层的页面调用 tesseract OCR（需另装 tesseract 与语言包） |
| `--ocr-lang L` | OCR 语言（默认 `chi_sim+eng`） |
| `--page-images` | 每页额外渲染一张整页 PNG 放在页首 |
| `--dry-run` | 只预览，不写入 |
| `--no-frontmatter` | 不写入 YAML frontmatter |

`obsidian_ppt2md.py` 支持 `--force` / `--dry-run` / `--no-frontmatter`。

## 转换规则

- 输出：`源文件名.md` + `attachments/`（提取的图片、整页渲染 PNG）；
- Markdown 内使用相对路径 `![](attachments/xxx.png)`，Obsidian 与通用渲染器均可显示；
- md 头部写入 YAML frontmatter（`source`、`converted_at`）；
- 已生成且未变更 → 自动跳过（增量）；
- 已存在同名 md 但非本插件生成 → 不覆盖（保护模式），`--force` 才覆盖。

## 已知限制

- 扫描件（纯图片 PDF）没有文本层，直接转换只能得到图片；需要 `--ocr`（另装 tesseract-ocr 与中文语言包）或使用「每页渲染为图片」；
- 复杂版式的文字顺序按位置排序提取，可能不够完美；
- PPT 的备注页、动画、母版内容不提取。
