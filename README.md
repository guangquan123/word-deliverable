# word-deliverable

按固定规范生成可直接正式交付的 Word（`.docx`）文档的 Skill。

## 这是什么

一个用来生成正式交付型 Word 文档的 Skill：输入 Markdown 或结构化内容，输出符合央国企、大型企业交付规范的 `.docx`。

产出目标是「打开即用」——封面、原生目录（字段已更新）、文档控制表、页眉页脚齐备，拿到手就能直接阅读、评审、打印、发送，不需要再人工调格式。

## 能力

- 四套排版 Profile：`government`（政企正式，默认）、`consulting`（咨询报告）、`technical`（技术方案）、`reader`（简洁阅读），成套参数，不混搭
- Word 原生 Heading 1-6 + 多级自动编号，两套编号体系：`formal`（第一章 / 1.1 / 1.1.1）、`lightweight`（一、/（一）/ 1. /（1））
- 原生 TOC 字段，生成后自动更新为真实标题与页码，不留「请按 F9 更新」
- 原生表格（表头浅灰底、重复标题行、≤7 列自动转横向 Section）、页眉页脚、修订记录表
- 生成后进行 DOCX/XML 结构校验，并可导出 PDF 做无图片文本坐标验收
- 默认禁图：不生成、不搜索、不调用图片模型

## 目录结构

```
SKILL.md                          Skill 入口与完整工作流
references/
  content-to-blocks.md            Markdown → body blocks 转换规则
  format-rules.md                 排版与格式规则全集
  validation-checklist.md         交付前验收清单
scripts/
  generate_docx.py                按 spec.json 生成 .docx
  update_fields.py                更新 TOC 与所有字段（Word COM / LibreOffice）
  validate_docx.py                结构、标题、目录、表格、分页验收检查
  render_pdf.py                   更新字段后导出 PDF 并做文本坐标验收
  render_pdf_uno.py               LibreOffice UNO 后端的 PDF 导出
  validate_pdf.py                 PDF 文本层、页面边界、空白页检查
tests/                            脚本单元测试
```

## 环境依赖

- Python 3
- `pip install python-docx pywin32`
  - `python-docx`：生成 .docx（必需）
  - `pywin32`：Word COM 调用，仅 Windows，用于更新字段、导出 PDF

更新目录字段还需以下之一：

| 优先级 | 工具 | 平台 | 能力 |
| --- | --- | --- | --- |
| 1 | Microsoft Word + pywin32 | Windows | TOC + 全字段更新，最佳 |
| 2 | LibreOffice UNO | 全平台 | TOC + 字段更新 |
| 3 | LibreOffice headless convert | 全平台 | 降级方案，TOC 可能不更新 |

含多级编号的文档建议用 `update_fields.py --libreoffice`，避免 Word COM 规范化 `numbering.xml` 时清空非标准编号级别。

## 安装

把本仓库内容放入 OpenCode / DeepWorks 的 skills 目录：

```
<项目根>/.opencode/skills/word-deliverable/
```

## 用法

1. 按 `SKILL.md` 的 JSON Schema 写出 `spec.json`（参数 + `body` block 列表）
2. 生成文档

```bash
python scripts/generate_docx.py spec.json output.docx
```

3. 更新目录与所有字段

```bash
python scripts/update_fields.py output.docx --libreoffice
```

4. 验收

```bash
python scripts/validate_docx.py output.docx spec.json
python scripts/render_pdf.py output.docx output.pdf spec.json
```

生成、更新、验收三步必须全部通过才能交付。

## 许可

[MIT](LICENSE)
