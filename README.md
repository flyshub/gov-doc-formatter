# 公文正文排版技能 (Gov-Doc-Formatter)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python: 3.9+](https://img.shields.io/badge/Python-3.9%2B-blue.svg)](https://www.python.org/)

Claude Code 技能：将纯文本或 Word 文档（`.docx` / `.doc`）按 **GB/T 9704-2012《党政机关公文格式》** 进行正文排版，生成 A4 标准公文 `.docx` 文件。

## 功能特性

- 📄 **标准页面**：A4 纸，天头 37mm / 订口 28mm / 下 28mm / 右 26mm，行距固定 28 磅
- 🔤 **四级层级标题**：自动识别 `一、`（黑体）→ `（一）`（楷体）→ `1.`/`1、`（仿宋）→ `(1)`/`（1）`（仿宋）
- 📝 **正文排版**：仿宋_GB2312 三号，首行缩进 2 字符，英文/数字 Times New Roman
- 🖼️ **图片保留**：自动提取原文图片（含 EMF 矢量图转 PNG），居中插入对应位置
- 📊 **表格保留**：自动提取表格，表头黑体/数据仿宋五号
- 🔢 **自动编号解析**：Word 的自动编号（`一、` / `（一）` 等）不丢
- 📑 **页码**：页脚 `— 1 —` 宋体四号居中
- 📦 **批量处理**：支持 `.docx` / `.doc` / `.txt` 混合批量输入
- 🌐 **全角半角兼容**：括号 `（一）` `(一)` 和序号 `1.` `1、` 均识别

## 快速开始

### 安装依赖

```bash
pip install python-docx pillow lxml olefile matplotlib

# .doc 旧格式需要（仅 Windows）
pip install pywin32
```

### 单个文件

```bash
# .docx 文件
python scripts/extract_docx.py 输入.docx /tmp/images

# .doc 旧格式（需要 Windows + MS Word）
python scripts/extract_doc.py 输入.doc /tmp/images

# 排版生成
python scripts/format_body.py --input data.json --output 公文.docx
```

### 批量处理

```bash
python scripts/batch_format.py file1.docx file2.doc file3.txt -o ./output/
```

### Claude Code 中使用

将此技能安装到 `~/.claude/skills/gov-doc-formatter/`，然后对话中直接说：

> "排成公文格式"  
> "把这个文档排成红头文件"

## 输入格式支持

| 输入格式 | 提取方式 | 编号 | 图片 | 表格 |
|---------|---------|:--:|:--:|:--:|
| `.docx` | `extract_docx.py`（python-docx + lxml） | ✅ | ✅ | ✅ |
| `.doc`（旧） | `extract_doc.py`（Word COM → .docx → 解析） | ✅ | ✅ | ✅ |
| 纯文本 | 直接按行拆分 | 手动 | — | — |

## 层级识别规则

| 层级 | 序号格式 | 正则 | 字体 | 字号 |
|------|---------|------|------|------|
| 标题 | — | — | 方正小标宋简体 | 二号 22pt |
| 一级 | `一、` | `^[一二三四五…]+、` | 黑体 | 三号 16pt |
| 二级 | `（一）` `(一)` | `^[（(][一二三四五…]+[）)]` | 楷体_GB2312 | 三号 16pt |
| 三级 | `1.` `1、` | `^\d+[.、]` | 仿宋_GB2312 | 三号 16pt |
| 四级 | `(1)` `（1）` | `^[（(]\d+[）)]` | 仿宋_GB2312 | 三号 16pt |
| 正文 | — | — | 仿宋_GB2312 | 三号 16pt |

## 文件结构

```
gov-doc-formatter/
├── SKILL.md                  # Claude Code 技能定义
├── REFERENCE.md              # GB/T 9704 参数手册
├── README.md                 # 本文件
└── scripts/
    ├── format_body.py        # 核心排版引擎
    ├── extract_docx.py       # .docx 提取器
    ├── extract_doc.py        # .doc 旧格式提取器
    └── batch_format.py       # 批量处理脚本
```

## 局限

- 仅排版正文和附件，不处理红头部分（发文机关标志、发文字号等由办公室后期添加）
- 不生成 PDF（仅输出 .docx）
- 旧格式 `.doc` 需要 Windows + MS Word + pywin32
- 方正小标宋简体需单独安装（商用授权），未安装时回退为宋体

## 开发历程

开发过程中解决了 16 个关键问题：

| 类别 | 典型问题 |
|------|---------|
| 编码 | .doc GBK 乱码 / 中文弯引号 JSON 损坏 / Windows 终端 GBK↔UTF-8 |
| 字体 | 中文名 vs 系统英文名不匹配 → 别名映射 + 回退链条 |
| 编号 | Word 自动编号在 numbering.xml → 解析 numPr |
| 图片 | EMF 矢量图不兼容 → PIL 转 PNG |
| 表格 | 遍历段落丢失表格 → body 子元素顺序遍历 |
| 架构 | 170行 god-function → ELEMENTS 数据驱动管道 |

## License

MIT
