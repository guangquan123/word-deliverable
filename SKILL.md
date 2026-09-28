---
name: word-deliverable
description: 按固定规范生成可直接正式交付的 Word（.docx）文档。当用户提供正文内容或文档生成任务，并希望输出符合央国企/大型企业正式项目交付规范的 Word 文件（实施方案、需求规格说明书、调研报告、数据治理方案、标准规范、项目报告、技术方案等）时使用。支持政企正式、咨询报告、技术方案和简洁阅读四套排版 Profile，默认禁图，提供智能表格、局部横向页、原生 TOC/页码字段、DOCX/XML 校验与 PDF 文本坐标验收。
metadata:
  author: "guangquan123"
---

# Word 正式交付文档生成器

## 何时使用

当用户要求生成一份**可直接向客户、领导或项目组正式交付的 Word 文档（.docx）**时使用本 skill。典型触发：

- "帮我生成一份 Word 实施方案 / 需求规格说明书 / 调研报告 / 数据治理方案 / 标准规范 / 项目报告 / 技术方案"
- "把这份内容整理成正式 Word 文档"
- "按规范出一份 .docx"
- "生成可交付的 Word"
- 用户提供正文或任务描述，并希望得到 `.docx` 而非 Markdown/HTML/PDF

**不用于**：只要 Markdown/HTML/PDF、PPT、海报、非正式内部速记、纯格式说明文档。

## 设计原则

本 skill 继承一套固定的 Word 视觉与工程规范（黑白灰正式风格），不随每次任务重新设计。核心是"继承格式，生成内容"。优先级：

1. 本 skill 固化的格式规则（最高）
2. 用户针对本次文档的特殊格式要求
3. AI 默认行为（最低，禁止自行重新设计）

最终目标：**用户打开 `.docx` 后即可直接阅读、评审、编辑、打印、发送，不需要再人工调整任何格式。**

## 环境前置与降级策略

本 skill 的"目录已生成、字段已更新、打开即用"承诺强依赖外部工具更新 Word 字段，**必须在使用前确认环境**。

### Python 依赖

```
pip install python-docx pywin32
```

- `python-docx`：生成 .docx（必需，所有平台）
- `pywin32`：Word COM 调用（仅 Windows，用于更新字段/导出 PDF）

### 字段更新依赖（至少需其一）

| 优先级 | 工具 | 平台 | 能力 | 说明 |
|--------|------|------|------|------|
| 1 | Microsoft Word + pywin32 | Windows | TOC + 全字段更新 | 最佳，目录/页码完全可靠 |
| 2 | LibreOffice UNO | 全平台 | TOC + 字段更新 | 需 LibreOffice 自带 python.exe（Windows 路径 `C:\Program Files\LibreOffice\program\python.exe`） |
| 3 | LibreOffice headless convert-to | 全平台 | 仅基本字段，TOC 可能不更新 | 最后兜底，**TOC 可能仍为占位文本** |

LibreOffice 路径默认硬编码为 `C:\Program Files\LibreOffice\program\soffice.exe`；Mac/Linux 需确保 `soffice` 在 PATH 中或通过环境变量 `LO_PATH` 指定。

### 失败停止规则（硬性）

1. **`update_fields.py` 三种方式全部失败时，必须停止，不得交付**。此时 .docx 的目录仍是占位文本"（目录将在打开后自动生成）"，交付即失败。应向用户报告环境缺失并给出安装建议。
2. **`validate_docx.py` 报告任何"问题清单"项时，必须修复后重跑"生成 → 更新字段 → 验收"全流程**，不得跳过验收直接交付。
3. **`validate_docx.py` 自动检查通过后，默认采用无图片验收**：可用 `render_pdf.py` 导出临时 PDF，并直接检查 PDF 的页数、文本层、文本坐标、页面边界和结构信息；**不得默认把 PDF 页面转成 PNG/JPG，也不得调用图片生成模型**。只有用户明确要求“视觉复检、逐页看图、截图检查”，或无图片检查发现无法判断的高风险版式异常且用户同意时，才可在 `.deepworks/tmp/` 临时栅格化少量疑似页面。
4. 若环境只支持 LibreOffice convert-to（优先级 3），TOC 可能无法更新——此时应明确告知用户"目录需在 Word/WPS 中打开后按 F9 更新"，**这视为降级交付，不算满足"打开即用"硬规则**。

## 工作流

### 第 1 步：收集文档参数

向用户确认或从上下文提取以下参数。缺失项用合理默认值并告知用户：

| 参数 | 说明 | 默认值 |
|------|------|--------|
| project_name | 项目名称 | 必填 |
| doc_name | 文档名称（未设则回退到 document_title） | 必填 |
| document_title | 文档标题（从 Markdown 首 `#` 识别的文件名，不进正文章节编号；未设 doc_name 时作封面文档名兜底） | 无 |
| subtitle | 副标题（可选） | 无 |
| doc_type | 文档类型 | 实施方案 |
| version | 版本 | V1.0 |
| org | 编制单位 | 项目组 |
| date | 编制日期 YYYY-MM-DD | 今天 |
| author | 编制人 | 编制单位 |
| status | 文档状态 | 正式发布 |
| layout_profile | 排版 Profile：`government` / `consulting` / `technical` / `reader` | government（政企正式） |
| chapter_num_format | 章编号模式：`chinese`=第一章+1.1 / `arabic`=第1章+1.1（仅 formal numbering profile 生效） | chinese |
| numbering_profile | 编号 Profile：`formal`=第一章/1.1/1.1.1/1.1.1.1/（1）/1） / `lightweight`=一、/（一）/1./（1）/1)/① | formal |
| heading_align_left | 标题是否全部靠左（无逐级缩进）：`true` 时 H1-H6 编号与正文都从左边距开始。用户明确要求"标题靠左/不要逐级缩进"时设为 true | false |
| toc_levels | 目录层级：`1-3`(默认) / `1-4`(规格说明书) / `1-5` / `1-6`(用户明确要求完整目录) | 1-3 |
| cover / toc / doc_control / header_footer | 是否需要封面/目录/文档控制/页眉页脚 | 全部是 |
| include_figures | 是否允许把用户提供的图片写入正文；不负责生成图片 | false（仅用户明确要求时为 true） |
| visual_review | 验收方式：`none`=仅结构校验 / `pdf-text`=PDF 无图片校验 / `raster`=页面图片复检 | pdf-text |

### 第 2 步：识别文档标题 + 建立标题树 + 规划正文结构（block 列表）

**2a. 识别文档标题**：若用户提供的正文是 Markdown，最前面的 `# 标题` 可能只是**文档名称**（如"# XX平台建设实施方案"）而非正文章节。判定规则：位于全文最前、全文唯一总标题、内容明显属文件名、后续标题均从 `##` 或更深层开始 → 识别为 `document_title`，**不参与正文章节自动编号**，可填入 spec 的 `document_title` 字段（同时可作封面文档名兜底）。

**2b. 相对层级映射**：识别文档标题后，剩余正文最浅层标题映射为 H1，逐层下推。**不要机械执行 `#=H1 / ##=H2`**，而是按内容的相对层级映射。例如 `## 建设方法 → H1`、`### 需求调研 → H2`、`#### 访谈安排 → H3`。

**2c. 建立标题树**：规划 block 列表前，先在内部建立完整标题树，确认父子关系合理后才开始生成。不得边读边临时猜标题样式。

**2d. 规划 block 列表**：根据用户提供的正文或任务描述，规划成统一的 block 列表。**严格采用用户提供的结构；用户未提供时按文档类型合理规划。** block 类型见下方 JSON Schema。完整的 Markdown→body 转换规则（相对映射、跳级处理、手工编号去除、标题vs列表区分）见 `references/content-to-blocks.md`。

**2e. 内容感知排版预处理**：在不改变事实、口径和章节语义的前提下，先判断内容形态再选择原生 Word 结构：

- 连续步骤、条件或职责清单 → 原生编号/项目符号列表，不堆成超长正文段落；
- 多对象同维度比较、字段定义、责任分工 → 表格；若维度过多则拆表，不制造超宽表；
- 原则、结论、注意事项 → `note` 或小节标题 + 正文，不使用彩色卡片或装饰图片；
- 单段过长时，仅按语义拆段；不得改写事实、补造结论或为了版面强行碎片化；
- 标题树通常控制在 3～4 个常用层级；H5/H6 仅在原文确有深层结构时使用；
- 默认不创建、不生成、不搜索任何图片。`figure` 仅用于用户明确提供并要求保留/插入的图片。

### 第 3 步：组装 JSON 并调用生成脚本

将参数 + 正文 block 列表写入一个 JSON 文件（建议放 `outputs/<组名>/_work/spec.json`，避免污染产物面板），然后调用：

```bash
python .opencode/skills/word-deliverable/scripts/generate_docx.py <spec.json> <output.docx>
```

脚本用 python-docx 生成完整 .docx（页面、样式、封面、目录 TOC 字段、文档控制、正文、页眉页脚、智能表格，以及仅在明确启用时插入的图片）。

### 第 4 步：更新 TOC 与所有字段（关键，不可跳过）

```bash
python .opencode/skills/word-deliverable/scripts/update_fields.py <output.docx>
# 含多级编号的文档推荐用 LibreOffice 后端（避免 Word COM 破坏编号）：
python .opencode/skills/word-deliverable/scripts/update_fields.py <output.docx> --libreoffice
```

脚本通过 Word COM（本机 Microsoft Word）打开文档，更新目录、PAGE、NUMPAGES 等所有字段，重新分页后保存。**这是满足"目录已真实生成、不让用户更新"硬规则的必要步骤。** 若 Word COM 不可用，自动回退到 LibreOffice headless 更新。

**已知坑（含多级编号的文档必读）**：Word COM 更新字段时会"规范化" numbering.xml，可能清空多级编号的非标准级别（如 `isLgl=True` 的 ilvl 1-8），导致 H2/H3/H4 自动编号消失。含多级编号的文档应使用 `--libreoffice`（或 `WD_NO_WORD=1`）强制走 LibreOffice UNO 后端，它不破坏编号定义，且保存后自动修复非首 section 的页眉页脚链接（移除独立 reference 使其继承前节）。UNO 路径完成后无需额外后处理脚本。

### 第 5 步：验收检查

```bash
python .opencode/skills/word-deliverable/scripts/validate_docx.py <output.docx> [spec.json]
```

脚本检查结构、标题（H1-H6 统计 + 跳级检测 + 手工编号残留 + H5/H6 误转列表 + 文档标题误编号）、目录、字段、表格、分页、页眉页脚等，输出检查报告。传入 spec.json 可按开关检查 TOC、页眉页脚和文档标题。默认导出 PDF 并执行无图片文本坐标验收：

```bash
python .opencode/skills/word-deliverable/scripts/render_pdf.py <output.docx> <output.pdf> [spec.json]
```

也可单独复核既有 PDF：

```bash
python .opencode/skills/word-deliverable/scripts/validate_pdf.py <output.pdf>
```

`visual_review=none` 仅做 DOCX/XML 检查；`pdf-text`（默认）检查文本层、页面边界、空白页、页尾孤立标题和末页残余；`raster` 只表示用户明确要求人工图片复检，本脚本不会自动栅格化。

### 第 6 步：修复与交付

验收发现的问题必须修复后重新生成/更新，**不得把任何格式修复工作留给用户**。通过验收后，将最终 `.docx` 放在 `outputs/` 下（如 `outputs/实施方案.docx`），并告知用户文件路径。

## 输入 JSON Schema

```json
{
  "layout_profile": "government",
  "include_figures": false,
  "visual_review": "pdf-text",
  "project_name": "XX集团信息化建设项目",
  "doc_name": "总体实施方案",
  "subtitle": "一期基础平台与二期业务应用建设",
  "doc_type": "实施方案",
  "version": "V1.0",
  "org": "XXX项目组",
  "date": "2026-09-01",
  "author": "张三",
  "status": "正式发布",
  "cover": true,
  "toc": true,
  "doc_control": true,
  "header_footer": true,
  "header_text": "XX集团信息化建设项目总体实施方案",
  "revision_history": [
    {"version": "V1.0", "date": "2026-09-01", "content": "初版编制", "author": "张三", "reviewer": "李四", "status": "正式发布"}
  ],
  "body": [
    {"type": "h1", "text": "项目概述"},
    {"type": "h2", "text": "项目背景"},
    {"type": "h3", "text": "业务现状"},
    {"type": "p", "text": "普通正文段落……"},
    {"type": "list", "ordered": true, "items": ["第一项", "第二项"]},
    {"type": "note", "title": "建设原则", "text": "局部组织用小节标题，不进目录。"},
    {"type": "table", "caption": "能力分层规则", "headers": ["层级", "名称", "说明"], "rows": [["A1", "基础层", "一期建设范围"]], "widths": [15, 25, 60]},
    {"type": "figure", "caption": "系统建设总体路径", "image": "outputs/xxx/diagram.png"}
  ]
}
```

### block 类型说明

| type | 说明 | 必填字段 |
|------|------|----------|
| h1 | 一级标题（第一章 XXX，自动编号） | text |
| h2 | 二级标题（1.1 XXX，自动编号） | text |
| h3 | 三级标题（1.1.1 XXX，自动编号） | text |
| h4 | 四级标题（1.1.1.1 XXX，自动编号） | text |
| h5 | 五级标题（（1） XXX，局部编号，每个 H4 下重新编号；Word 内部仍为 Heading 5） | text |
| h6 | 六级标题（1） XXX，局部编号，每个 H5 下重新编号；Word 内部仍为 Heading 6） | text |
| p | 正文段落 | text |
| list | 列表 | items[]; ordered(布尔) |
| note | 小节标题 + 说明（不进目录） | title, text |
| table | 表格 | caption, headers[](非空), rows[][]; widths[](可选,正数权重), landscape(可选) |
| figure | 图片；仅 `include_figures=true` 时插入 | caption, image（相对 spec.json 所在目录或绝对路径） |
| pagebreak | 强制分页 | 无 |

> **标题层级规则**：内容层级决定 Heading，Heading 决定 Word 结构，多级列表只负责显示编号。H5/H6 虽显示为局部编号（（1）/1）），但 Word 内部仍是 Heading 5/6 Style，**不得降级为 Normal 或普通列表**。Markdown→body 的相对层级映射、跳级处理、手工编号去除等规则见 `references/content-to-blocks.md`。

## 核心格式规则（摘要）

完整规则见 `references/format-rules.md`，验收清单见 `references/validation-checklist.md`。

- **视觉风格**：黑白灰正式，央国企/大型企业交付文档风格。禁止 AI 默认蓝色标题、彩色色块、渐变、卡片、装饰图标、阴影、霓虹色。
- **页面与正文**：由 `layout_profile` 成套控制；默认 `government` 为 A4 纵向、页边距上 3.7 / 下 3.5 / 左 2.8 / 右 2.6 cm、仿宋 14pt、固定行距 28pt。`consulting`、`technical`、`reader` 使用各自成套参数；禁止跨 Profile 混搭。
- **文档结构**：封面 → 目录 → 文档控制 → （前言）→ 正文 → 附录。顺序不得随意调整。
- **封面**：独立成页，大留白，项目名 36pt 黑体加粗居中，文档名 36pt 黑体加粗居中，封面信息表（版本/编制单位/编制日期）居中、左列浅灰底加粗、黑色细边框，禁止跨页。
- **目录**：Word 原生 TOC 字段，点线引导符，页码右对齐。**必须已更新显示真实标题、层级、页码**，不得出现"请按 F9 更新"等提示。
- **标题**：Word Heading Style + 多级自动编号（H1-H6 全 6 级），禁止手工输入章节编号。
  - H1：第一章 XXX（黑体 22pt 加粗，`chinese` 模式中文计数 / `arabic` 模式第1章）
  - H2：1.1 XXX（黑体 16pt 加粗）
  - H3：1.1.1 XXX（黑体 14pt 加粗）
  - H4：1.1.1.1 XXX（黑体 14pt 加粗）
  - H5：（1） XXX（正文字号 14pt 加粗，**局部编号**，每个 H4 下重新编号；Word 内部仍是 Heading 5，不降级 Normal/列表）
  - H6：1） XXX（正文字号 14pt 加粗，**局部编号**，每个 H5 下重新编号；Word 内部仍是 Heading 6，不降级 Normal/列表）
  - **章编号模式**：`chapter_num_format: "chinese"`（默认，第一章 + 1.1/1.1.1）或 `"arabic"`（第1章 + 1.1/1.1.1），通过 `isLgl`(legal format) 标志让中文一级与阿拉伯二三级共存。仅 `formal` 编号 Profile 生效。
  - **编号 Profile**：`numbering_profile: "formal"`（默认，上述编号体系）/ `"lightweight"`（一、/（一）/1./（1）/1)/①，适用于会议纪要/工作记录等轻量文档）。无论哪种 Profile，**Heading 层级本身不能丢失**。
  - 标题与后续内容同页（Keep with next）。父级变化后下级编号自动重置（Word 自动编号负责）。
- **目录层级**：`toc_levels` 控制 TOC 显示深度，默认 `1-3`（正式长篇方案）；规格说明书/标准规范可 `1-4`；用户明确要求完整目录可 `1-5`/`1-6`。H5/H6 默认不进目录但保留在 Word 导航结构与自动编号体系。原则：标题结构可以深，但目录不能因此不可阅读。
- **正文（默认 `government`）**：仿宋 14pt 黑色，两端对齐，首行缩进 2 字（约 28pt），固定行距 28pt，段前段后 0；其他排版 Profile 使用其对应 Style 参数。
- **小节标题**：不进目录的局部小标题用独立"小节标题"Style，禁止用正文+手工加粗代替。
- **表格**：Word 原生 Table，固定布局，黑色细边框，表头浅灰底加粗居中，正文垂直居中，短字段居中/长文本左对齐。`headers` 必须非空，所有行列数一致；`widths` 如提供须与列数一致、均为正数，脚本按比例归一化。列宽按当前 Section 可用宽度和内容权重计算；`landscape=true` 或列数 `>=7` 时只为该表建立横向 Section，表后恢复纵向并继承页眉页脚。重复标题行，禁止拆行和孤立尾行。
- **图片**：默认完全禁用，不生成、不搜索、不调用图片模型，也不写错误占位。只有 `include_figures=true` 才处理用户提供的 `figure`；路径为空、不存在或加载失败必须停止并报告可操作错误。
- **页眉**：正文页统一显示"项目名称+文档名称"，黑色 10.5pt 左对齐，下方黑色细横线；封面不显示页眉页脚。
- **页脚**：居中"第 X 页 共 Y 页"，用 PAGE + NUMPAGES 自动字段，页码连续，禁止手工页码。
- **样式驱动**：全部用 Word Style（正文/Heading 1-6/小节标题/列表正文/表题/图题/表头/表格正文/说明注释），禁止 Normal+手工加粗模拟标题、禁止大面积直接格式。
- **分页**：用 Page Break / Section Break / Keep with next / Keep lines together 控制，禁止连续回车/空行/空格顶页。非设计性空白页 = 0，无孤行孤页。
- **零过程文字**：正式文档不得出现"请按 F9/请更新域/TODO/待生成/示例占位符/AI说明/请用户自行调整"等制作过程文字。

## 脚本说明

| 脚本 | 作用 |
|------|------|
| `scripts/generate_docx.py` | 用 python-docx 按 spec.json 生成完整 .docx |
| `scripts/update_fields.py` | 用 Word COM / LibreOffice 更新 TOC 与所有字段并保存 |
| `scripts/validate_docx.py` | 结构/标题/目录/字段/表格/分页/页眉页脚验收检查 |
| `scripts/render_pdf.py` | 更新字段后可靠导出新 PDF，并按 spec 执行默认文本坐标验收 |
| `scripts/validate_pdf.py` | 无图片检查 PDF 文本层、页面边界、空白页和异常稀疏分页 |

## 交付要求

- 最终交付物为可编辑 `.docx`，放在 `outputs/` 下。
- 中间 spec.json、检查报告等放 `outputs/<组名>/_work/`，不污染产物面板。
- 交付前必须跑完"生成 → 更新字段 → 验收"，验收通过才能交付。
- 交付时告知用户 `.docx` 的精确工作区相对路径。
