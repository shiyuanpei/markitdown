# Office2Markdown 开发者文档

## 项目概述

Office2Markdown 通过修改 markitdown 和 mammoth 库实现高级功能。

## 架构设计

### 补丁系统

所有修改保存在项目目录，通过 `install_patches.py` 安装到系统库。

**markitdown补丁** (`markitdown_patches/`):
- `converters/_docx_converter.py` - 相对路径 + LLM OCR
- `converters/_pptx_converter.py` - 相对路径 + LLM描述

**mammoth补丁** (`mammoth_patches/`):
- `docx/body_xml.py` - OMML公式转换 (+150行)
- `docx/office_xml.py` - 元素注册

### 关键修改

#### 1. 相对路径支持

**markitdown/_docx_converter.py** (line 111):
```python
# 原版
return {"src": str(filepath.resolve())}

# 修改后
return {"src": f"{self.output_dir.name}/{filename}"}
```

**markitdown/_pptx_converter.py** (line 281):
```python
# 原版
filename = str(saved_path.resolve())

# 修改后
filename = f"{images_path.name}/{saved_path.name}"
```

#### 2. OMML公式转换

**mammoth/docx/body_xml.py** (~line 373):
```python
def omath(element):
    """Handle OMML equations"""
    from docxlatex.parser import OMMLParser
    
    xml_string = _xml_element_to_string(element)
    etree_element = fromstring(xml_string)
    
    parser = OMMLParser()
    latex = parser.parse(etree_element)
    
    return _success(documents.text(f"${latex}$"))
```

## 维护指南

### 更新补丁

1. 修改系统库文件:
```bash
vim "C:\Users\...\markitdown\converters\_docx_converter.py"
```

2. 测试:
```bash
python office2md.py test.docx
```

3. 保存补丁:
```bash
cp "C:\Users\...\converters\_docx_converter.py" markitdown_patches/converters/
```

4. 更新文档:
```bash
vim markitdown_patches/README.md
```

### 库更新后

```bash
pip install --upgrade markitdown mammoth
python install_patches.py
python office2md.py test.docx  # 测试
```

### 打包发布

```bash
# 清理旧版本
rm -rf build dist

# 打包
python -m PyInstaller office2md.spec --clean --noconfirm

# 测试exe
dist/office2md.exe test.docx

# 检查文件大小
ls -lh dist/office2md.exe
```

## 测试

### 单元测试

```bash
# 测试DOCX转换
python office2md.py docDemo1.docx

# 测试PPTX转换  
python office2md.py pptDemo1.pptx

# 测试批量转换
python office2md.py *.docx --no-llm
```

### 验证输出

检查项:
- [ ] Markdown语法正确
- [ ] 图片路径是相对路径
- [ ] OMML公式转换为LaTeX
- [ ] 图片成功提取

## 技术栈

### 核心依赖

- **markitdown**: 文档转换引擎
- **mammoth**: DOCX解析
- **docxlatex**: OMML→LaTeX
- **python-pptx**: PPTX处理
- **Pillow**: 图像处理
- **openai**: LLM客户端
- **magika**: 文件类型检测

### PyInstaller配置

**office2md.spec** 关键配置:
```python
datas = [
    (magika_models, 'magika/models'),
    (magika_config, 'magika/config'),
]

hiddenimports = [
    'markitdown',
    'mammoth',
    'docxlatex',
    'magika',
    'onnxruntime',
]
```

## 已知问题

### 已解决
- [x] 图片绝对路径 → 相对路径
- [x] OMML不转换 → 添加docxlatex
- [x] PPT无描述 → 添加LLM
- [x] Unicode编码错误 → 使用ASCII

### 待优化
- [ ] PyInstaller打包时间长 (~3分钟)
- [ ] EXE文件较大 (~158MB)

## 版本历史

- **v1.0.3** - 完整补丁系统，文档整理
- **v1.0.2** - 修复Unicode编码，清理项目
- **v1.0.1** - 添加mammoth补丁
- **v1.0.0** - 初始版本

## 代码规范

- 使用Python 3.12+
- 遵循PEP 8
- 补丁文件保持与原库一致的代码风格
- 关键修改添加注释标记
