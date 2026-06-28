# word-solution-doc

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/)
[![python-docx](https://img.shields.io/badge/python--docx-1.0+-green.svg)](https://python-docx.readthedocs.io/)
[![GitHub stars](https://img.shields.io/github/stars/yushengxin/word-solution-doc.svg)](https://github.com/yushengxin/word-solution-doc/stargazers)
[![GitHub release](https://img.shields.io/github/v/release/yushengxin/word-solution-doc.svg)](https://github.com/yushengxin/word-solution-doc/releases)

> 生成符合中文解决方案排版规范的 Word 文档（仿宋小四正文、黑体多级标题、首页+目录+正文结构、自动页码）。

📄 配套的 OpenClaw / Claude Code Skill，使用 `python-docx` 把 Markdown 一键转成符合国内"方案/可研/投标"文档排版规范的 `.docx`。

---

## ✨ 特性

- **规范排版** — 仿宋小四正文，黑体多级标题（五级），1.5 倍行距，首行缩进 2 字符
- **多级编号** — 一级用中文数字（一、二、三），二级用阿拉伯数字（1、2、3），三级以下 1.1 / 1.1.1 / 1.1.1.1
- **首页 + 目录 + 正文** — 标准三段式结构，页码从目录页开始计算
- **自动页码** — 页脚底部居中，Word 域实现
- **表格规范** — 自动调整、居中、首行加粗
- **Markdown 输入** — 简单 Markdown 即可生成，不写复杂模板

---

## 📦 安装

### 依赖

- Python 3.8+
- `python-docx >= 1.0.0`
- `lxml >= 4.6.0`

```bash
pip install -r requirements.txt
```

### 作为 OpenClaw Skill 使用

将整个 `word-solution-doc` 文件夹复制到你的 OpenClaw workspace 的 `skills/` 目录下即可。

```
~/.openclaw/workspace/skills/word-solution-doc/
```

---

## 🚀 快速上手

### 1. 写一个 Markdown 文件

`content.md`：

```markdown
# 项目概述

本项目旨在构建一个企业级文档协作平台，支持多人实时编辑、版本管理、权限控制等核心能力。

## 项目背景

随着公司业务规模扩大，团队协作场景日益复杂，传统 Office 工具难以满足多人协作需求。

## 项目目标

本项目目标是在六个月内完成平台一期建设，覆盖文档协作、权限管理、版本控制三大核心模块。

# 技术方案

## 架构设计

整体架构分为四层：接入层、网关层、业务服务层、数据层。

| 模块 | 技术栈 | 说明 |
| --- | --- | --- |
| 用户中心 | Spring Boot | 用户管理 |
| 文档服务 | MongoDB | 文档存储 |

### 总体架构

接入层使用 Nginx 负责静态资源分发与负载均衡。
```

### 2. 一行命令生成

```bash
python scripts/generate.py \
  --input content.md \
  --output 解决方案.docx \
  --title "企业级文档协作平台解决方案" \
  --subtitle "—— 一期建设方案"
```

### 3. 不想写 Markdown？直接生成示例

```bash
python scripts/generate.py --demo --output example.docx
```

---

## 📐 排版规范

| 项目 | 规范 |
|------|------|
| 页面 | A4 纵向，页边距上下 2.54cm / 左右 3.17cm |
| 正文 | 仿宋 小四（12pt），1.5 倍行距，首行缩进 2 字符 |
| 一级标题 | 黑体 二号（22pt） — 序号格式：一、二、三 |
| 二级标题 | 黑体 小二（18pt） — 序号格式：1、2、3 |
| 三级标题 | 黑体 三号（16pt） — 序号格式：1.1、1.2 |
| 四级标题 | 黑体 小三（15pt） — 序号格式：1.1.1 |
| 五级标题 | 黑体 四号（14pt） — 序号格式：1.1.1.1 |
| 表格 | 仿宋 五号（10.5pt），居中，表头加粗 |
| 页码 | 目录页起算，页面底部居中 |
| 首页 | 无页码 |

**关键规则：**
- 一级标题的序号**不在**下级标题中继承（即二级从 1 重新开始）
- 各级标题**无**首行缩进
- 所有序号由 Word 多级列表自动管理

完整规范见 [`references/format-spec.md`](references/format-spec.md)。

---

## 🛠️ CLI 参数

```
python scripts/generate.py [选项]

选项:
  -o, --output    输出 .docx 路径（必填）
  -t, --title     文档主标题（默认：解决方案）
      --subtitle  文档副标题（可选）
  -i, --input     输入 Markdown 文件路径
      --demo      使用内置示例内容
```

---

## 📁 目录结构

```
word-solution-doc/
├── README.md                   # 本文件
├── LICENSE                     # MIT
├── requirements.txt            # Python 依赖
├── .gitignore
├── SKILL.md                    # OpenClaw Skill 描述
├── references/
│   └── format-spec.md          # 完整排版规范
├── scripts/
│   └── generate.py             # 核心生成脚本
├── examples/
│   └── example-output.docx     # 示例输出
└── assets/                     # 图片等资源
```

---

## 💡 在 AI 对话中使用

把项目放好之后，直接告诉 AI：

> "按 word-solution-doc 规范生成 Word 文档"

AI 会自动调用本 skill 的脚本生成符合规范的 `.docx`。

---

## 📜 License

MIT — 详见 [LICENSE](LICENSE)
