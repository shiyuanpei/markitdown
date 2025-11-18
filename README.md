# markitdown - 增强版

本仓库是 [markitdown](https://github.com/microsoft/markitdown) 的 fork 版本，专门针对中文技术文档和数学公式的转换进行了增强。

## 新增功能概览

### 1. 数学公式增强

#### 1.1 行间公式自动空行
为行间公式（`$$...$$`）后自动添加两个空行，防止公式在 Markdown 中粘连。

**问题**：
```markdown
$$公式1$$
$$公式2$$
$$公式3$$
```
公式会连在一起显示。

**解决方案**：
```markdown
$$公式1$$


$$公式2$$


$$公式3$$
```

#### 1.2 OMML Base64 解码
解码来自 mammoth 的 Base64 编码的 OMML 公式占位符，还原为正确的 LaTeX 代码。

**处理流程**：
```
⟨OMML:$:YmFzZTY0X2RhdGE=⟩  →  $\LaTeX代码$
```

#### 1.3 MathType 公式支持

支持 Word 文档中的 **MathType** 和 **Equation Editor** 公式。

**公式类型**：
- **内置公式编辑器** (Microsoft Equation) - 使用 OMML 格式存储
- **MathType** - 第三方公式编辑器，可能使用不同格式

**处理机制**：
1. OMML 格式公式：通过 docxlatex 转换为 LaTeX
2. MathType 公式：如果可识别，尝试提取 LaTeX 表示
3. 图像形式的公式：提取为图片文件

**兼容性**：
- 新版 Word (2007+) 的内置公式 - 完全支持
- MathType 6.x/7.x - 部分支持
- 旧版 Equation Editor 3.0 - 基本支持

### 2. 图像处理增强

#### 2.1 WMF/EMF 自动转 PNG
自动检测并转换 WMF/EMF 格式图像为 PNG，因为 Markdown 不支持 WMF/EMF 显示。

**技术特性**：
- 使用 ImageMagick 转换
- 高质量输出（600 DPI）
- 白色背景，RGB 色彩空间
- 转换成功后自动删除原始 WMF/EMF 文件
- 转换失败时优雅降级，保留原文件

**转换参数**：
```bash
magick -density 600 input.wmf \
       -background white \
       -alpha remove \
       -colorspace RGB \
       -quality 100 \
       output.png
```

#### 2.2 LLM 图像描述生成

集成 LLM (大语言模型) 能力，自动为图像生成 Markdown 格式的 alt 文本描述。

**功能特性**：
- 自动识别图像内容
- 生成详细的中文描述
- 支持技术图表、流程图、示意图等
- 可选功能，需要配置 LLM API

**使用场景**：
```markdown
![LLM自动生成的图像描述](image_001.png)
```

**配置要求**：
- OpenAI API 或兼容的 LLM 服务
- 设置环境变量：`OPENAI_API_KEY`

**LLM 提示词模板**：
```
描述这张图片的内容，使用简洁的中文。
重点关注：技术细节、数据关系、流程步骤。
```

**隐私说明**：
- LLM 描述功能默认关闭
- 图像数据发送到配置的 LLM 服务
- 建议在处理敏感文档时禁用此功能



### 3. PPTX 处理增强

#### 3.1 幻灯片结构保留

将 PowerPoint 幻灯片转换为 Markdown 时保持层次结构。

**转换规则**：
- 每张幻灯片 → 二级标题 (`##`)
- 幻灯片编号自动添加
- 标题和正文内容分离
- 列表和表格结构保留

**示例输出**：
```markdown
## 幻灯片 1: 项目概述

项目背景介绍...

## 幻灯片 2: 技术架构

- 前端框架
- 后端服务
- 数据库设计
```

#### 3.2 PPTX 图像提取

自动提取幻灯片中的图片、图表和 SmartArt。

**提取内容**：
- 插入的图片文件
- 图表 (Chart) - 转换为图片
- SmartArt - 转换为图片
- 形状组合 - 转换为图片

**命名规则**：
```
演示文稿_images/
  ├── slide_01_image_001.png
  ├── slide_01_image_002.png
  ├── slide_02_chart_001.png
  └── slide_03_smartart_001.png
```

#### 3.3 PPTX 中的公式

支持 PowerPoint 中插入的数学公式。

**公式来源**：
- 内置公式编辑器 - 转换为 LaTeX
- MathType 公式 - 尝试提取 LaTeX
- 公式图片 - 保留为图像

**转换示例**：
```markdown
## 幻灯片 5: 控制方程

雷诺方程：

$$\frac{\partial p}{\partial \theta} + \frac{\partial p}{\partial r} = 0$$
```

### 4. 集成 office2md 工具

提供简化的命令行界面，专门用于 Office 文档到 Markdown 的转换。

**特性**：
- 自动检测文件类型（DOCX, PPTX, XLSX）
- 自动创建图像输出目录
- 批量处理支持
- 详细的转换日志

**使用方式**：
```bash
# 转换单个文件
office2md document.docx

# 转换多个文件
office2md file1.docx file2.pptx file3.xlsx
```

### 5. 依赖增强

集成了增强版的依赖库：
- **docxlatex** - Unicode 符号自动转换
- **python-mammoth** - OMML Base64 编码

## 修改的文件

### converters/_docx_converter.py

1. **decode_omml_placeholder()** 函数：
   - Base64 解码 OMML 占位符
   - 为行间公式添加空行

2. **DocxImageWriter.__call__()** 方法：
   - WMF/EMF 检测和转换逻辑
   - ImageMagick 集成

### converters/_pptx_converter.py

1. **幻灯片结构处理**：
   - 保留幻灯片编号和标题层次
   - 正文内容格式化

2. **图像和图表提取**：
   - SmartArt 转图片
   - Chart 图表处理
   - 形状组合渲染

### 新增文件

- `src/markitdown/office2md.py` - office2md 命令行工具
- `DEVELOPER.md` - 开发者文档
- `README_OFFICE2MD.md` - office2md 使用指南
- `office2md.spec` - PyInstaller 打包配置

### pyproject.toml

更新依赖配置：
```toml
[project.optional-dependencies]
all = [
  "git+https://github.com/shiyuanpei/docxlatex.git@main",
  # ... 其他依赖
]
docx = [
  "mammoth~=1.11.0",
  "lxml",
  "git+https://github.com/shiyuanpei/docxlatex.git@main"
]

[project.scripts]
markitdown = "markitdown.__main__:main"
office2md = "markitdown.office2md:main"
```

## 完整转换流程

```
Word DOCX
    ↓
[mammoth] 读取 DOCX，提取 OMML
    ↓
[docxlatex] OMML → LaTeX (Unicode符号转换)
    ↓
[mammoth] LaTeX → Base64 编码 → 占位符
    ↓
[mammoth] 生成 HTML (包含占位符)
    ↓
[markdownify] HTML → Markdown (占位符保持不变)
    ↓
[markitdown] 解码占位符 → LaTeX
[markitdown] WMF/EMF → PNG
[markitdown] 添加公式空行
    ↓
最终 Markdown 输出
```

## 安装

从 GitHub 安装：

```bash
pip install git+https://github.com/shiyuanpei/markitdown.git@main
```

安装后会同时提供 `markitdown` 和 `office2md` 两个命令。

## 使用示例

### 基本使用

```bash
# 使用 office2md（推荐）
office2md technical_paper.docx

# 使用 markitdown
markitdown technical_paper.docx -o output.md
```

### 批量转换

```bash
office2md *.docx
```

### 包含数学公式的文档

输入 (Word):
```
雷诺方程：

∂p/∂θ + ∂p/∂r = 0

其中 p 为压力，θ 为角度。
```

输出 (Markdown):
```markdown
雷诺方程：

$$\partial p/\partial \theta + \partial p/\partial r = 0$$


其中 p 为压力，θ 为角度。
```

## 技术文档应用场景

本增强版特别适合：
- 机械工程技术文档
- 流体力学论文
- 数学物理教材
- 包含大量公式和图像的科技文档
- 中文学术文档

## 与原项目的关系

- 保持与上游 markitdown 的兼容性
- 所有增强功能都是非侵入式添加
- 可以作为 markitdown 的直接替代品使用

## 相关项目

- [docxlatex 增强版](https://github.com/shiyuanpei/docxlatex) - Unicode 符号映射
- [python-mammoth 增强版](https://github.com/shiyuanpei/python-mammoth) - OMML Base64 编码

## 原始项目

原始 markitdown 项目: https://github.com/microsoft/markitdown

## 许可证

与原项目保持一致的 MIT 许可证。
