# Word 解决方案文档 — 完整排版规范

## 1. 页面设置

- 纸张大小：A4（210mm × 297mm）
- 页边距：上下 2.54cm，左右 3.17cm（默认值，可按需调整）
- 文档方向：纵向

## 2. 字体映射

| 中文字号 | 磅值(pt) | 用途 |
|---------|---------|------|
| 二号 | 22pt | 一级标题 |
| 小二 | 18pt | 二级标题 |
| 三号 | 16pt | 三级标题 |
| 小三 | 15pt | 四级标题 |
| 四号 | 14pt | 五级标题 |
| 小四 | 12pt | 正文 |
| 五号 | 10.5pt | 表格文字 |

- **黑体**：所有标题
- **仿宋**：正文、表格文字

## 3. 段落规范

### 正文段落

- 字体：仿宋，小四（12pt）
- 首行缩进：2 字符（约 480 twips ≈ 24pt）
- 行距：1.5 倍行距
- 段前段后间距：0

### 标题段落

- 字体：黑体（见上表字号）
- 首行缩进：无
- 行距：1.5 倍行距
- 段前段后间距：0

## 4. 标题序号规则

### 多级列表格式

```
一、一级标题
1、二级标题
1.1 三级标题
1.1.1 四级标题
1.1.1.1 五级标题
```

### 序号关键规则

1. **一级标题用中文大写数字**，后跟顿号"、"（如：一、二、三）
2. **二级标题用阿拉伯数字**，后跟顿号"、"（如：1、2、3）
3. **三级及以下用点号分隔**，无顿号（如：1.1、1.1.1、1.1.1.1）
4. **一级序号不继承**：二级标题重新从 1 开始编号，不显示一级序号
5. 所有序号由多级列表自动管理，确保正确递增

### 示例

```
一、项目概述
1、项目背景
1.1 业务需求
1.1.1 核心需求
1.1.1.1 功能需求
1.1.1.2 非功能需求
1.1.2 补充需求
1.2 技术现状
2、项目目标
二、技术方案
1、架构设计
1.1 总体架构
```

## 5. 表格规范

- 自动调整：根据窗口自动调整（`AUTOFIT_WINDOW`）
- 对齐：所有单元格文字水平居中、垂直居中
- 标题行：首行文字**加粗**
- 字体：仿宋，五号（10.5pt）
- 边框：默认边框（单线）

## 6. 页面结构

### 首页

- 单独占一页
- 内容通常包含：文档标题、版本号、日期、编制单位等
- 无页码
- 首页后插入分页符

### 目录

- 单独占一页
- 使用 python-docx 的 TOC 字段自动生成
- **页码从目录页开始计算**（目录页 = 第 1 页）
- 目录后插入分页符

### 正文

- 紧跟目录之后
- 页码续接目录页码

## 7. 页码

- 首页：无页码
- 目录页：页码起始为 1
- 正文：续接目录页码
- 页码位置：页面底部居中
- 页码格式：阿拉伯数字（1, 2, 3…）

## 8. python-docx 实现要点

### 字体设置

```python
from docx.shared import Pt
from docx.oxml.ns import qn

def set_font(run, font_name, size_pt, bold=False):
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.name = font_name
    # 设置中文字体
    r = run._element
    rPr = r.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        r.insert(0, rPr)
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), font_name)
```

### 行距设置

```python
from docx.shared import Twips
from docx.enum.text import WD_LINE_SPACING

def set_line_spacing_1_5(paragraph):
    pf = paragraph.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = 1.5
```

### 首行缩进

```python
def set_first_line_indent(paragraph, chars=2, font_size_pt=12):
    # 2 字符 ≈ font_size_pt * 2 的 twips
    paragraph.paragraph_format.first_line_indent = Pt(font_size_pt * chars)
```

### 多级列表

python-docx 原生不支持多级列表，需要通过 XML 操作实现。核心思路：
1. 定义多级列表的 AbstractNum 和 Num
2. 为每个级别设置编号格式和缩进
3. 在段落上应用对应的列表级别

详见 `scripts/generate.py` 中的实现。
