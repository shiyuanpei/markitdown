# markitdown - 增强版

本仓库是 [markitdown](https://github.com/microsoft/markitdown) 的 fork 版本,专门针对包含数学公式和图片的技术文档转换进行了深度增强。

## 核心改进

### 与原版 markitdown 的对比

| 功能 | 原版 markitdown | 本增强版 |
|------|----------------|----------|
| 文字提取 | ✅ | ✅ |
| Word 内置公式 (OMML) | ❌ 不支持 | ✅ 完整支持 |
| MathType 公式 | ❌ 不支持 | ✅ 完整支持 |
| 图片提取 | ❌ 不支持 | ✅ 完整支持 |
| WMF/EMF 转换 | ❌ | ✅ 自动转换 |
| Unicode 符号 (docxlatex) | ❌ 不转换 | ✅ 80+ 符号映射 |
| LLM 图片描述 | ⚠️ 基础支持 | ✅ 增强支持 |
| 中文文档 | ❌ | ✅ 完整中文文档 |


## 本增强版的核心功能

### 1. 完整的数学公式支持 ⭐

**将 Word/PPT 中的数学公式完整转换为 LaTeX 格式**,支持:

- **内置公式编辑器** (Microsoft Equation) - OMML 格式,完全支持
- **MathType 公式** - 第三方公式编辑器,部分支持
- **行内公式** `$...$` 和 **行间公式** `$$...$$` 自动识别
- **Unicode 数学符号自动转换** - ∂→\partial, θ→\theta, ∇→\nabla 等 80+ 符号


### 2. 智能图片处理 ⭐

**自动提取并处理文档中的所有图片**,包括:

- **图片提取** - 自动保存到独立目录
- **格式转换** - WMF/EMF 自动转换为 PNG(Markdown 兼容)
- **LLM 图片描述** - 集成大语言模型,自动生成图片的详细中文描述(可选)
- **高质量输出** - 600 DPI,保证技术图表清晰度


## 安装方法

### 前置要求

**ImageMagick(必需)** - 用于 WMF/EMF → PNG 格式转换
- Windows: https://imagemagick.org/script/download.php#windows
- Linux: `sudo apt-get install imagemagick` 或 `sudo yum install imagemagick`
- macOS: `brew install imagemagick`

**LLM API(可选)** - 用于图片智能描述
- 推荐使用 OpenRouter: https://openrouter.ai/
- 或 OpenAI API

### 从 GitHub 安装

确保已安装 ImageMagick 后,运行:

```bash
pip install -r https://github.com/shiyuanpei/markitdown/raw/main/requirements.txt
```

这一条命令会自动安装增强版的 docxlatex、python-mammoth 和 markitdown(包含 office2md 工具)。

## 使用方法

### 命令行工具

#### office2md(推荐,简化版)

```bash
# 转换单个文件
office2md document.docx

# 批量转换
office2md *.docx *.pptx

# 输出到指定目录
office2md report.docx -o output/
```

#### markitdown(完整版,更多选项)

```bash
# 基本转换
markitdown document.docx -o output.md

# 启用 LLM 图片描述
markitdown slides.pptx -o slides.md --llm-client openai

# 转换并查看
markitdown paper.docx | less
```

### Python API

```python
from markitdown import MarkItDown

# 基本使用
md = MarkItDown()
result = md.convert("technical_report.docx")
print(result.text_content)

# 启用 LLM 图片描述
from openai import OpenAI
md = MarkItDown(llm_client=OpenAI(), llm_model="gpt-4-vision-preview")
result = md.convert("presentation.pptx")
```

## 支持的文件格式

- **DOCX** - Word 文档(公式 + 图片完整支持)
- **PPTX** - PowerPoint 演示文稿(公式 + 图片完整支持)
- **XLSX** - Excel 表格
- **PDF** - PDF 文档(文本提取)
- **图片** - PNG, JPG, GIF 等(LLM 描述)
- **音频** - MP3, WAV 等(转录)
- **其他** - HTML, ZIP, EPUB 等

## 技术架构

### 依赖关系

本项目通过 fork 和增强三个核心库实现功能:

1. **docxlatex** - Unicode 符号到 LaTeX 的映射
   - 仓库: https://github.com/shiyuanpei/docxlatex
   - 功能: 80+ Unicode 数学符号自动转换

2. **python-mammoth** - OMML 公式 Base64 编码保护
   - 仓库: https://github.com/shiyuanpei/python-mammoth
   - 功能: 防止 LaTeX 代码被 HTML 转义破坏

3. **markitdown**(本项目) - 整合转换和图片处理
   - 仓库: https://github.com/shiyuanpei/markitdown
   - 功能: 统一转换接口 + 图片处理 + LLM 集成

### 转换流程

```
Word/PPT 文档
    ↓
[python-mammoth 增强版]
  - 读取 OMML 公式
  - Base64 编码保护 LaTeX
    ↓
[docxlatex 增强版]
  - OMML → LaTeX 转换
  - Unicode 符号映射
    ↓
[markitdown 增强版]
  - 解码 LaTeX 公式
  - WMF/EMF → PNG 转换
  - LLM 图片描述(可选)
  - 格式优化
    ↓
完整的 Markdown 输出
  - 公式 ✓
  - 图片 ✓
  - 文字 ✓
```

## 限制和已知问题

- **MathType 兼容性**: MathType 6.x/7.x 部分支持,旧版可能无法完全转换
- **复杂公式**: 极其复杂的嵌套公式可能需要手动调整
- **LLM 成本**: 图片描述功能需要调用 API,会产生费用
- **ImageMagick 依赖**: WMF/EMF 转换需要安装 ImageMagick

## 贡献和反馈

欢迎提交 Issue 和 Pull Request!

- 问题反馈: https://github.com/shiyuanpei/markitdown/issues
- 功能建议: 欢迎讨论新的增强方向

## 相关项目

- [docxlatex 增强版](https://github.com/shiyuanpei/docxlatex) - Unicode 符号映射
- [python-mammoth 增强版](https://github.com/shiyuanpei/python-mammoth) - OMML Base64 编码
- [原版 markitdown](https://github.com/microsoft/markitdown) - Microsoft 官方版本

## 许可证

与原项目保持一致的 MIT 许可证。

---

**注**: 本项目保持与上游 markitdown 的兼容性,可作为原版的直接替代品使用。
