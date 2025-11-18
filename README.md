# markitdown - 增强版

本仓库是 [markitdown](https://github.com/microsoft/markitdown) 的 fork 版本，**专门针对包含数学公式和图片的技术文档转换进行了深度增强**。

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

**将 Word/PPT 中的数学公式完整转换为 LaTeX 格式**，支持：

- **内置公式编辑器** (Microsoft Equation) - OMML 格式，完全支持
- **MathType 公式** - 第三方公式编辑器，部分支持
- **行内公式** `$...$` 和 **行间公式** `$$...$$` 自动识别
- **Unicode 数学符号自动转换** - ∂→\partial, θ→\theta, ∇→\nabla 等 80+ 符号

**转换示例**：

Word 文档中的公式：
```
雷诺方程：∂p/∂θ + ∂p/∂r = 0
```

转换后的 Markdown：
```markdown
雷诺方程：

$$\partial p/\partial \theta + \partial p/\partial r = 0$$
```

### 2. 智能图片处理 ⭐

**自动提取并处理文档中的所有图片**，包括：

- **图片提取** - 自动保存到独立目录
- **格式转换** - WMF/EMF 自动转换为 PNG（Markdown 兼容）
- **LLM 图片描述** - 集成大语言模型，自动生成图片的详细中文描述（可选）
- **高质量输出** - 600 DPI，保证技术图表清晰度

**LLM 描述示例**：

原始图片 → LLM 自动生成描述：
```markdown
![流体动力学润滑膜压力分布三维曲面图，显示压力在轴承中心区域达到峰值](image_001.png)
```

这使得 Markdown 文档对视障用户友好，也便于后续的文档检索和分析。

## 安装方法

### 前置要求

### ImageMagick（必需）

本工具依赖 **ImageMagick** 进行图片格式转换（WMF/EMF → PNG）。

**Windows 安装**：
1. 下载：https://imagemagick.org/script/download.php#windows
2. 选择推荐的安装包（如 `ImageMagick-7.1.1-Q16-HDRI-x64-dll.exe`）
3. 安装时**务必勾选** "Add application directory to your system path"
4. 安装后验证：打开命令行输入 `magick -version`

**Linux 安装**：

```bash
# Ubuntu/Debian
sudo apt-get install imagemagick

# CentOS/RHEL
sudo yum install imagemagick
```

**macOS 安装**：
```bash
brew install imagemagick
```

⚠️ **重要**：如果不安装 ImageMagick，WMF/EMF 图片将无法转换为 PNG，可能导致图片在 Markdown 中无法显示。

### 2. LLM API（可选，用于图片智能描述）

如果需要使用 LLM 自动生成图片描述功能，需要配置以下服务之一：

#### 推荐：OpenRouter（支持多种模型）

OpenRouter 是一个 LLM 聚合服务，支持 GPT-4 Vision、Claude 3 等多种视觉模型。

**配置步骤**：
1. 注册：https://openrouter.ai/
2. 获取 API Key
3. 设置环境变量：
   ```bash
   export OPENAI_API_KEY=your_openrouter_key
   export OPENAI_BASE_URL=https://openrouter.ai/api/v1
   ```

**优势**：
- 一个 API 访问多种模型
- 按需付费，价格透明
- 支持国内访问
- 兼容 OpenAI API 格式

#### 使用 OpenAI 官方

```bash
export OPENAI_API_KEY=your_openai_key
# 不需要设置 BASE_URL
```

#### 使用其他兼容服务

任何兼容 OpenAI API 格式的服务都可以使用：

```bash
export OPENAI_API_KEY=your_api_key
export OPENAI_BASE_URL=https://your-service-url/v1
```

💡 **提示**：LLM 图片描述是可选功能，不配置也可以正常转换文档，只是图片不会有自动生成的描述文字。

## 从 GitHub 安装（推荐）

**确保已安装 ImageMagick 后**，按顺序运行以下命令：

```bash
# 1. 安装增强版 docxlatex（Unicode 符号映射）
pip install git+https://github.com/shiyuanpei/docxlatex.git@main

# 2. 安装增强版 python-mammoth（OMML Base64 编码保护）
pip install git+https://github.com/shiyuanpei/python-mammoth.git@master

# 3. 安装增强版 markitdown（包含 office2md 工具）
pip install "git+https://github.com/shiyuanpei/markitdown.git@main#subdirectory=packages/markitdown"
```

⚠️ **重要**：必须按照上述顺序安装，以确保增强版的 docxlatex 和 python-mammoth 被正确使用。

### 依赖关系

本项目通过 fork 和增强三个核心库实现功能：

1. **docxlatex** - Unicode 符号到 LaTeX 的映射
   - 仓库：https://github.com/shiyuanpei/docxlatex
   - 功能：80+ Unicode 数学符号自动转换

2. **python-mammoth** - OMML 公式 Base64 编码保护
   - 仓库：https://github.com/shiyuanpei/python-mammoth
   - 功能：防止 LaTeX 代码被 HTML 转义破坏

3. **markitdown** (本项目) - 整合转换和图片处理
   - 仓库：https://github.com/shiyuanpei/markitdown
   - 功能：统一转换接口 + 图片处理 + LLM 集成

### 额外依赖

**必需**：
- ImageMagick - 用于 WMF/EMF 到 PNG 的转换
  - Windows: https://imagemagick.org/script/download.php
  - 安装后确保 `magick` 命令可用

**可选**（用于 LLM 图片描述）：
- OpenAI API 或兼容的 LLM 服务（推荐使用 **OpenRouter** 聚合服务）
- 环境变量配置：
  ```bash
  # 使用 OpenAI
  export OPENAI_API_KEY=your_openai_key

  # 使用 OpenRouter（推荐，支持多种模型）
  export OPENAI_API_KEY=your_openrouter_key
  export OPENAI_BASE_URL=https://openrouter.ai/api/v1
  ```
- OpenRouter 注册：https://openrouter.ai/

## 使用方法

### 命令行工具

安装后提供两个命令：

#### 1. office2md（推荐，简化版）

```bash
# 转换单个文件
office2md document.docx

# 批量转换
office2md *.docx *.pptx

# 输出到指定目录
office2md report.docx -o output/
```

#### 2. markitdown（完整版，更多选项）

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

- **DOCX** - Word 文档（公式 + 图片完整支持）
- **PPTX** - PowerPoint 演示文稿（公式 + 图片完整支持）
- **XLSX** - Excel 表格
- **PDF** - PDF 文档（文本提取）
- **图片** - PNG, JPG, GIF 等（LLM 描述）
- **音频** - MP3, WAV 等（转录）
- **其他** - HTML, ZIP, EPUB 等

## 典型应用场景

本增强版特别适合：

✅ **机械工程技术文档** - 大量公式和示意图
✅ **流体力学/传热学论文** - 偏微分方程和数值结果
✅ **数学物理教材** - 复杂公式和理论推导
✅ **科研报告** - 实验数据图表和分析
✅ **学术论文** - 中英文混合，公式密集
✅ **工程手册** - 技术规范和计算公式

## 详细功能说明

### 数学公式增强

#### OMML 到 LaTeX 转换
- 通过 docxlatex 将 Office MathML 转换为标准 LaTeX
- 支持分数、根式、上下标、积分、求和、矩阵等

#### Unicode 符号映射
- 80+ 常用数学符号自动转换
- 希腊字母：α, β, γ, δ, θ, λ, μ, ω, Γ, Δ, Σ, Ω...
- 运算符：∂, ∇, ∞, ∑, ∏, ±, ×, ÷, ·...
- 关系符：≈, ≠, ≡, ≤, ≥, ∈, ⊂, ⊃...

#### 公式格式优化
- 行间公式自动添加空行，防止粘连
- 行内公式和行间公式自动识别
- LaTeX 代码完整性保护（Base64 编码）

### 图片处理增强

#### 自动图片提取
- DOCX/PPTX 中的所有图片自动导出
- 按文档结构组织（如按幻灯片编号）
- 支持嵌入图片和链接图片

#### WMF/EMF 格式转换
- 自动检测 Windows 元文件格式
- 使用 ImageMagick 转换为 PNG
- 高质量输出：600 DPI, RGB 色彩空间
- 转换失败时优雅降级

#### LLM 图片描述（可选）
- 集成 OpenAI GPT-4 Vision 或兼容模型
- 自动生成详细中文描述
- 识别技术图表、流程图、数据图等
- 可自定义提示词模板

### PPTX 特殊处理

#### 幻灯片结构保留
- 每张幻灯片转为二级标题
- 自动添加幻灯片编号
- 保持标题和正文层次

#### 图表和 SmartArt
- Chart 图表转为图片
- SmartArt 图形转为图片
- 形状组合渲染为图片

## 打包为可执行文件

提供 PyInstaller 配置文件，可打包为独立 .exe 文件：

```bash
cd d:/python/markitdown
pyinstaller office2md.spec
```

生成的 `dist/office2md.exe` 可在没有 Python 环境的 Windows 机器上直接运行。

## 技术架构

### 增强的转换流程

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
  - LLM 图片描述（可选）
  - 格式优化
    ↓
完整的 Markdown 输出
  - 公式 ✓
  - 图片 ✓
  - 文字 ✓
```

## 技术细节

### 修改的核心文件

#### converters/_docx_converter.py
- `decode_omml_placeholder()` - Base64 解码 OMML 占位符
- `DocxImageWriter.__call__()` - WMF/EMF 转换和图片处理

#### converters/_pptx_converter.py
- 幻灯片结构处理
- SmartArt 和 Chart 图表提取

#### src/markitdown/office2md.py
- 简化的命令行接口
- 批量处理支持

## 限制和已知问题

- **MathType 兼容性**：MathType 6.x/7.x 部分支持，旧版可能无法完全转换
- **复杂公式**：极其复杂的嵌套公式可能需要手动调整
- **LLM 成本**：图片描述功能需要调用 API，会产生费用
- **ImageMagick 依赖**：WMF/EMF 转换需要安装 ImageMagick

## 贡献和反馈

欢迎提交 Issue 和 Pull Request！

- 问题反馈：https://github.com/shiyuanpei/markitdown/issues
- 功能建议：欢迎讨论新的增强方向

## 相关项目

- [docxlatex 增强版](https://github.com/shiyuanpei/docxlatex) - Unicode 符号映射
- [python-mammoth 增强版](https://github.com/shiyuanpei/python-mammoth) - OMML Base64 编码
- [原版 markitdown](https://github.com/microsoft/markitdown) - Microsoft 官方版本

## 许可证

与原项目保持一致的 MIT 许可证。

---

**注**：本项目保持与上游 markitdown 的兼容性，可作为原版的直接替代品使用。
