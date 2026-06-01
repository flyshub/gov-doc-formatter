---
name: gov-doc-formatter
description: 将纯文本或 Word 文档按公文正文格式排版。自动识别四级层级标题（一、/（一）/1./(1)）并应用对应字体字号（黑体/楷体/仿宋），正文仿宋三号，行距固定值28磅，英文数字 Times New Roman。生成 A4 标准 .docx。Use when 用户说"排成公文"、"公文格式"、"红头文件"、"正文排版"、"公文排版"。
---

# 公文正文排版

## Quick Start

用户粘贴文本或上传 Word，说"排成公文格式"即可。

## 流程

### 1. 输入预处理

- **`.docx`** → 使用 `extract_docx.py` 提取文本（**必须用此脚本**，解析 Word 自动编号）
  ```bash
  python <skill_dir>/scripts/extract_docx.py <输入文件>
  ```
- **`.doc`**（旧格式）→ 使用 `extract_doc.py`，通过 Word COM 转为 `.docx` 后再提取编号
  ```bash
  python <skill_dir>/scripts/extract_doc.py <输入文件>
  ```
  > 需要 Windows + MS Word + `pywin32` 包。Word 需已安装。
- **纯文本** → 按换行拆分为段落数组，跳过空行
- **确认**：展示提取的段落列表（含编号），问用户是否继续

### 2. 要素提取与确认

提取并确认以下要素：
- **标题**：如原文无标题，AI 生成并标注 `[AI自动生成，请核实]`
- **称谓**：行文对象（顶格，如"北京致远互联软件股份有限公司："），可选
- **正文**：主体内容，自动识别四级层级（`一、`/`（一）`/`1.`/`(1)`）
- **附件**：附件列表，可选
- **署名**：行文单位（右对齐），可选
- **日期**：行文日期（右对齐，位于署名正下方），可选

**输出顺序**：标题 → 称谓 → 正文 → 附件 → 署名 → 日期

**强制展示确认**，未确认不进入下一步。

### 3. 确认输出

将所有提取的要素以列表形式展示给用户确认。用户确认"准确"/"生成"后进入代码生成阶段。未确认不继续。

### 4. 代码生成

将确认后的内容构建为 JSON，通过 **stdin 管道**传给脚本（推荐）或写入临时文件。

JSON 结构：
```json
{
  "标题": "...",
  "称谓": "...",
  "正文": [
    "普通段落（仿宋三号，自动识别层级标题）",
    {"text": "自定义对齐段落", "align": "right|center|left", "indent": false},
    {"image": "/tmp/img_001.png", "width_mm": 120, "height_mm": 80},
    ...
  ],
  "附件": ["附件1", "附件2", ...],
  "署名": "...",
  "日期": "2026年4月22日"
}
```

正文项支持三种格式：
- **字符串**：自动识别层级（`一、`/`（一）`/`1.`/`(1)`）并应用对应字体
- **文本对象**：`{"text": "...", "align": "right|center|left", "indent": false}` — 自定义对齐和缩进
- **图片对象**：`{"image": "/path/to/img.png", "width_mm": 120, "height_mm": 80}` — 居中插入，超出版心宽度等比缩放

**推荐：stdin 管道模式**（避免临时文件和 shell 转义问题）：
```bash
# 1. 用 Python 写 JSON 到工作目录（禁止 bash heredoc，中文引号会被破坏）
python -c "
import json
data = {...}  # 构建 JSON
with open('<工作目录>/_gov_input.json', 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False)
"

# 2. 管道传入脚本
PYTHONIOENCODING=utf-8 python <skill_dir>/scripts/format_body.py \
  --stdin \
  --output <工作目录>/<文件名>.docx \
  < <工作目录>/_gov_input.json

# 3. 清理临时文件
rm <工作目录>/_gov_input.json
```

> ⚠️ **禁止用 bash heredoc 构建 JSON**（中文弯引号 `""` 会被破坏）。必须用 Python 构建。
> ⚠️ **Windows 下避免用 `/tmp/`** 传临时文件——Git Bash 和 Python subprocess 的 `/tmp/` 路径映射可能不一致。

备选：文件模式（`--input`）：
```bash
PYTHONIOENCODING=utf-8 python <skill_dir>/scripts/format_body.py \
  --input <绝对路径>/gov_body_input.json \
  --output <工作目录>/<文件名>.docx
```

脚本 stdout 返回 JSON：`{"status": "ok", "path": "...", "warnings": [...]}`

### 5. 输出

- 告知生成路径
- 如有字体缺失警告，列出并建议安装对应字体
- 提醒用户核实层级序号、字体显示效果

## 约束

- 仅排版正文和附件，不处理发文机关标志、发文字号、签发人、主送/抄送机关
- 不生成 PDF，仅输出 .docx
- 不打包字体文件

## 参见

排版参数细节见 [REFERENCE.md](REFERENCE.md)
