---
name: paper-formatter
description: 学术论文 Word 格式调整与格式规范参考（GB/T 7713.2—2022）。两大模块：(1) 在用户上传的 Word 文档上以修订模式（Track Changes）自动调整格式——inspect_docx.py 体检 + LLM 元素分类映射 + revise_format.py 执行，全部改动记为可接受/拒绝的 Word 修订；(2) 中英论文格式规范库——字体字号表（表B.1）、章节/图表/公式编号、量和单位、数字用法、参考文献等。当用户要求调整论文格式、按国标排版、统一字体字号、用修订模式改 Word 格式时使用。
---

# Paper Formatter — 论文格式修订模式调整 + GB/T 7713.2 格式规范

处理用户上传的 Word 论文文档：按 GB/T 7713.2—2022《学术论文编写规则》做格式体检与修订模式调整。
两大能力模块：

- **模块A · 格式规范库**：`references/format-spec-cn-en.md`（清洗自 GB/T 7713.2—2022 原文，中英文元素均覆盖，
  应/宜/可 强度已保留）+ `scripts/format_rules.json`（机读版：表B.1 字号字体、字号→磅换算、每个元素的
  Word 实操参数与 detect_hints）。
- **模块B · 修订模式格式调整**：`scripts/inspect_docx.py`（只读体检）+ `scripts/revise_format.py`
  （Word COM 修订模式下应用格式）。技术细节与降级方案见 `references/revision-techniques.md`。

## 工作流（模块B 标准四步）

1. **体检**：
   ```bash
   python "<skill目录>/scripts/inspect_docx.py" <论文.docx> --out inspect.json
   ```
   得到逐段格式清单（段号 p<N>、文本、字体字号、对齐缩进、is_A4、表格数）。
   先读 inspect.json，不要凭空猜测文档结构。
2. **分类映射**：对照 format-spec 与 format_rules.json 的 detect_hints，判断每段元素类型
   （title_cn / authors_cn / affiliation_cn / abstract_cn_head / abstract_cn_body / keywords_cn_* /
   title_en / abstract_en_* / front_misc / heading_chapter / heading_section / body / caption_fig /
   caption_tab / table_body / ack_* / ref_head / ref_body / appendix_*），写 mapping.json：
   ```json
   {
     "paragraphs": {
       "1": "title_cn",
       "4": {"lead_prefix": "摘要：", "lead": "abstract_cn_head", "element": "abstract_cn_body"},
       "6": "heading_chapter"
     },
     "tables": {"1": "table_body"}
   }
   ```
   - "摘要：×××"这类引题+内容混排段必须用 lead_prefix 拆分。
   - 难判定的段落宁可不映射，也不要错判后强改；可在映射前把疑问告知用户。
   - 期刊有专门投稿模板时，以期刊模板优先于本标准（标准本身注明表B.1"可参考"）。
3. **预演**：`python revise_format.py <论文.docx> --map mapping.json --dry-run`，核对计划改动。
4. **执行**：去掉 --dry-run。输出 revisions_total 与逐段报告；随后可复跑 inspect 或导出 PDF 验证。

## 硬性注意

- 文件须先关闭（正被 Word/WPS 打开会失败）；脚本新开独立 Word 实例，不影响用户已开窗口。
- 环境依赖：pywin32（`pip install pywin32`）+ 本机 Word；本机已实测可用（见 revision-techniques.md 第6节）。
- 不主动接受/拒绝修订——那是作者的权利；只产出修订，不静默覆盖原文。
- 无 Word 环境的降级方案（python-docx 直改 + 改动对照表，无修订痕迹）见 revision-techniques.md 第5节，
  使用前必须明确告知用户。
- 字号速查：小2号=18pt，4号=14pt，小4号=12pt，5号=10.5pt，小5号=9pt。
