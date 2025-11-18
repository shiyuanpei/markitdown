# Office2Markdown

将Office文档（DOCX, PPTX）转换为Markdown，支持OMML公式转换和LLM图像识别。

## 快速开始（3步）

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 安装补丁

```bash
python install_patches.py
```

### 3. 使用

```bash
python office2md.py document.docx
```

## 主要功能

- ✅ **OMML公式**: Word原生公式自动转换为LaTeX
- ✅ **图像公式**: WMF/EMF图像公式使用LLM OCR识别（可选）
- ✅ **图片提取**: 自动提取，使用相对路径
- ✅ **PPTX描述**: 自动生成LLM图片描述（可选）
- ✅ **批量处理**: 支持通配符 `*.docx` `*.pptx`

## LLM配置（可选）

配置后可获得图片描述和公式OCR功能：

```bash
# Windows
set OPENROUTER_API_KEY=sk-or-v1-xxx

# Linux/Mac
export OPENROUTER_API_KEY=sk-or-v1-xxx
```

获取密钥: https://openrouter.ai/

**成本**: 约$0.01-0.10/100页，基本功能完全免费

不使用LLM:
```bash
python office2md.py document.docx --no-llm
```

## 命令行参数

```
office2md [-h] [-o OUTPUT] [-d IMAGES_DIR] [--no-llm] [-q] [--version] input [input ...]

参数:
  input           输入文件(支持通配符)
  -o OUTPUT       输出文件
  -d IMAGES_DIR   图片目录(默认: <文件名>_images/)
  --no-llm        禁用LLM功能
  -q, --quiet     安静模式
  --version       显示版本
```

## 使用示例

### 学术论文转换

```bash
python office2md.py thesis.docx
```

输出:
- `thesis.md` - 包含LaTeX公式
- `thesis_images/` - 提取的图片

### 演示文稿转换

```bash
python office2md.py presentation.pptx
```

输出:
- `presentation.md` - 包含图片描述
- `presentation_images/` - 提取的图片

### 批量转换

```bash
python office2md.py *.docx --no-llm
```

## 补丁说明

本项目修改了markitdown和mammoth库以支持高级功能：

**markitdown补丁**:
- 相对路径（原版是绝对路径）
- LLM OCR和图片描述

**mammoth补丁**:
- OMML公式转LaTeX

详见: `markitdown_patches/README.md` 和 `mammoth_patches/README.md`

### 库更新后重新安装补丁

```bash
pip install --upgrade markitdown mammoth
python install_patches.py
```

## 打包为EXE

```bash
pip install pyinstaller
python -m PyInstaller office2md.spec --clean --noconfirm
```

生成: `dist/office2md.exe`

## 项目结构

```
office2markdown/
├── office2md.py              # 主程序
├── install_patches.py        # 补丁安装
├── requirements.txt          # 依赖
├── README.md                 # 本文件
├── DEVELOPER.md              # 开发者文档
├── markitdown_patches/       # markitdown补丁
└── mammoth_patches/          # mammoth补丁
```

## 常见问题

**Q: 为什么需要补丁?**
A: 添加相对路径、OMML公式、LLM功能支持

**Q: 不装补丁能用吗?**
A: 可以，但缺少高级功能，图片路径是绝对路径

**Q: 公式准确吗?**
A: OMML 100%准确，图像公式90%+

**Q: 会影响其他程序吗?**
A: 不会。可用 `python install_patches.py --restore` 恢复

## 技术栈

- Python 3.12
- markitdown (修改版)
- mammoth (修改版)
- docxlatex (OMML转换)
- OpenRouter API (Gemini 2.5 Flash)

## 许可

MIT License
