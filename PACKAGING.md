# 电子发票号码提取工具 — 跨平台开发与打包指南

本文说明如何在 **macOS** 上开发与测试，以及在 **Windows** 上将单文件脚本打包为独立 `.exe`。

---

## 一、macOS 开发与测试

### 1. 环境要求

- Python 3.9 或更高版本（建议 3.10+）
- macOS 自带 `tkinter`，一般无需额外安装 GUI 库

### 2. 安装依赖（Conda 环境 `py39`）

在项目目录下激活已有 Conda 环境并安装依赖（**不要在项目里再建 `.venv`**）：

```bash
cd /path/to/fapiaoNew
conda activate py39
pip install -r requirements.txt
```

`requirements.txt` 仅包含：

- `pymupdf` — PDF 文本提取（`import fitz`）
- `openpyxl` — Excel 写入

### 3. 运行程序

```bash
conda activate py39
python invoice_extractor.py
```

### 4. 功能自测建议

1. 准备一个文件夹，放入若干**标准电子发票 PDF**（仅当前目录一层，不含子文件夹）。
2. 在界面中选择 PDF 文件夹与 Excel 导出目录。
3. 点击「开始极速提取」，观察日志是否逐文件输出 `文件名 -> 发票号码`。
4. 在导出目录确认生成 `发票提取结果_YYYYMMDD_HHMMSS.xlsx`，表头为 `文件名`、`发票号码`。
5. 用损坏或加密 PDF 测试：日志应显示「跳过：文件异常」，程序不崩溃。

### 5. macOS 说明

- 首次运行若被 Gatekeeper 拦截，请在「系统设置 → 隐私与安全性」中允许。
- 本工具**不建议在 macOS 上用 PyInstaller 产出给 Windows 用的 exe**；Windows 可执行文件应在 Windows 本机构建。
- 请统一使用 Conda 环境 `py39` 开发与运行，避免在项目目录下创建额外的 `python -m venv`。
- 若 `py39` 中缺少 tkinter，可执行：`conda install -n py39 tk`。

---

## 二、Windows 打包为独立 exe

### 1. 为何必须在 Windows 上打包

PyInstaller 打包的是**当前操作系统**下的 Python 运行时与原生库。在 Mac 上打出来的产物无法在 Windows 上运行。请将 `invoice_extractor.py` 与 `requirements.txt` 拷贝到 Windows 电脑（U 盘、网盘、Git 均可）。

### 2. Windows 环境准备

1. 从 [python.org](https://www.python.org/downloads/windows/) 安装 Python 3.9+。
2. 安装时**勾选** “Add python.exe to PATH” 以及 **“tcl/tk and IDLE”**（否则 `tkinter` 不可用）。
3. 打开 **命令提示符** 或 **PowerShell**，进入项目目录：

```cmd
cd C:\path\to\fapiaoNew
```

### 3. 安装依赖与 PyInstaller

```cmd
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
pip install pyinstaller
```

### 4. 一键打包命令（推荐）

在项目目录、已激活虚拟环境下执行：

```cmd
pyinstaller -F -w --name InvoiceExtractor invoice_extractor.py
```

参数说明：

| 参数 | 含义 |
|------|------|
| `-F` | 单文件 exe（onefile） |
| `-w` | 无控制台窗口（GUI 程序） |
| `--name InvoiceExtractor` | 输出 exe 名称 |

### 5. 打包产物位置

- 可执行文件：`dist\InvoiceExtractor.exe`
- 构建缓存：`build\`（可删除）
- 规格文件：`InvoiceExtractor.spec`（二次打包时可编辑）

将 **`dist\InvoiceExtractor.exe`** 单独复制给其他 Windows 用户即可运行（无需安装 Python）。

### 6. 若 exe 启动报错（缺少模块）

部分环境需显式收集 PyMuPDF，可改用：

```cmd
pyinstaller -F -w --name InvoiceExtractor ^
  --hidden-import=fitz ^
  --collect-all pymupdf ^
  invoice_extractor.py
```

（PowerShell 中将 `^` 换为行末反引号 `` ` ``，或写成一行。）

### 7. 常见问题

| 现象 | 处理 |
|------|------|
| 双击 exe 无反应 | 去掉 `-w` 重新打包，在 cmd 中运行 exe 查看报错 |
| 杀毒软件误报 | onefile 会解压到临时目录，可加入白名单或对 exe 签名 |
| 首次启动较慢 | onefile 需解压，属正常现象 |
| 找不到 tkinter | 重装 Python 并勾选 tcl/tk |
| PDF 提取失败 | 确认发票为标准电子版（含文本层），非纯扫描图 |

### 8. 分发给最终用户

仅需提供 `InvoiceExtractor.exe`。用户双击后：

1. 选择含 PDF 的文件夹；
2. 选择 Excel 保存目录；
3. 点击「开始极速提取」。

---

## 三、项目文件清单

| 文件 | 说明 |
|------|------|
| `invoice_extractor.py` | 主程序（单文件） |
| `requirements.txt` | 运行时依赖 |
| `PACKAGING.md` | 本文档 |

---

## 四、技术栈速查

- GUI：`tkinter` / `ttk`
- PDF：`PyMuPDF`（`fitz`），仅第一页 `get_text("text")`
- Excel：`openpyxl`（不使用 pandas）
- 路径：全程 `pathlib.Path`，无硬编码斜杠
