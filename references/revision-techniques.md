# Word 修订模式（Track Changes）格式调整技术参考

> 适用环境：Windows + 本机安装 Microsoft Word（COM 自动化）。本机已验证：`Word.Application` COM 可用；
> pywin32 安装：`python -m pip install pywin32`。若只有 WPS：尝试 ProgID `KWPS.Application`
> （API 与 Word 基本兼容），不可用则退回"无修订标记"方案（见第 5 节）。

## 1. 核心原理

- Word COM（pywin32）打开文档后设置 `doc.TrackRevisions = True`，之后**一切修改都以修订形式记录**，
  作者在 Word 中逐条接受/拒绝，原文不会被静默覆盖。
- **格式改动是否记为修订**由 `doc.TrackFormatting = True` 控制（Word 2010+ 有此属性；设置代码需
  try/except 兜底）。已验证：字体/字号/对齐/缩进改动会产生 `Type=3`（wdRevisionProperty，格式修订）
  与段落属性类修订，审阅窗格显示为"设置了格式: …"气泡。
- 中文字体设 `Font.NameFarEast`，西文字体设 `Font.Name`；两者分开设置互不干扰。

## 2. 标准工作流（对应本 skill 两个脚本）

1. **体检**：`python inspect_docx.py 论文.docx --out inspect.json`（只读）。
   得到逐段清单：段号 p<N>、文本预览、当前字体/字号/对齐/缩进、是否在表格内、页面尺寸（含 is_A4）。
   段号与 revise_format.py 使用**同一 Word Paragraphs 枚举**（1-based，含表格内段落），可直接复用。
2. **分类映射**：对照 `references/format-spec-cn-en.md` 判断每段属于哪个元素（题名/作者/摘要/章标题/正文/图题/参考文献…），
   写 mapping.json：
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
   - 段落键写 `"12"` 或 `"p12"` 均可。
   - 同一段内"引题+内容"混排（如"摘要：×××"）用 lead_prefix：前缀套引题格式，其余套内容格式。
   - 元素 id 一览见 `scripts/format_rules.json` 的 element_formats（含 detect_hints 辅助判断）。
3. **预演**：`python revise_format.py 论文.docx --map mapping.json --dry-run`——核对每段计划改动。
4. **执行**：去掉 --dry-run 再跑。输出 `revisions_total` 与逐段 applied 报告。
   - 默认就地保存（修订可回退）；`--out 新文件.docx` 另存（目标已存在则拒绝覆盖）。
5. **复检**：再跑一次 inspect_docx.py 确认格式到位；`revisions_total` 应大于 0。
   也可导出 PDF 目视验证修订气泡：`doc.SaveAs2(path, FileFormat=17)`。

## 3. 环境与运行要点

- **文件不能同时被 Word/WPS 打开**，否则 `Documents.Open` 抛错（脚本会提示）。执行前确认关闭。
- 脚本用 `DispatchEx("Word.Application")` 新开独立 Word 实例（`Visible=False, DisplayAlerts=0`），
  不干扰用户已开的 Word 窗口；务必 try/finally `word.Quit()` 防止僵尸 WINWORD 进程。
- 中文路径对 COM 无障碍（不同于 FAISS 的纯 ASCII 限制）。
- Windows 控制台默认 GBK：打印含特殊字符的输出前先 `sys.stdout.reconfigure(encoding="utf-8")`
  （两个脚本均已内置）。
- 逐段 COM 调用较慢（几百段约十几秒），属正常；不要用 `for para in doc.Paragraphs` 边枚举边改，
  枚举中途改动会使集合不稳定——按索引 `Item(i)` 取用。

## 4. 常用 COM 片段库（文本级修订，超出字体表范围时用）

```python
import pythoncom, win32com.client as win32
pythoncom.CoInitialize()
word = win32.DispatchEx("Word.Application"); word.Visible = False; word.DisplayAlerts = 0
doc = word.Documents.Open(abs_path, AddToRecentFiles=False)
doc.TrackRevisions = True
try:
    # ① 查找替换（记为修订）：如"20℃"→"20 ℃"（数值与单位间留空隙）
    #    注意 Word 通配符：勾选 MatchWildcards 时 ^w 为空白、{,} 为重复
    fr = doc.Content.Find
    fr.ClearFormatting(); fr.Replacement.ClearFormatting()
    fr.Text = "20℃"; fr.Replacement.Text = "20 ℃"
    fr.Execute(Replace=1)  # wdReplaceAll=2 仅替换下一个; 1=wdReplaceOne

    # ② 给某段加批注（用于向作者提问，不算修订）
    para = doc.Paragraphs.Item(6)
    doc.Comments.Add(para.Range, "此节编号应为 2.1，请确认层级")

    # ③ 段落级格式修订（revise_format.py 已封装）
    rng = para.Range
    rng.Font.NameFarEast = "黑体"; rng.Font.Size = 12

    # ④ 修订统计
    n = doc.Revisions.Count
    doc.Save()
finally:
    doc.Close(SaveChanges=0); word.Quit(); pythoncom.CoUninitialize()
```

- wdAlignParagraph：left=0, center=1, right=2, justify=3。
- 混合格式的 Range 读属性返回 `9999999`（wdUndefined）——inspect 脚本已转成 "MIXED"。
- 不要主动 `AcceptAllRevisions()`/`RejectAllRevisions()`——接受与否是作者的权利；
  仅当用户明确要求"直接接受全部修订"时才执行。

## 5. 无 Word / 无 pywin32 的降级方案

- **python-docx 直改**（无修订标记）：直接改 run.font / paragraph_format，改完另存副本，
  并生成"改动对照表"（段号 | 原格式 | 新格式）供人工核对。**必须明确告知用户此路径没有修订痕迹。**
- 保留修订语义的高级做法（仅必要时）：解包 docx 直接操作 OOXML——文本增删用 `w:ins`/`w:del`
  （配 `w:author`/`w:date`），格式改动在 `w:pPr`/`w:rPr` 内嵌 `w:pPrChange`/`w:rPrChange` 保存旧值。
  复杂度高，除非用户明确要求"不装 Word 也要真修订"，否则不要走这条路。

## 6. 已验证结论（2026-09-19 本机实测）

- `Word.Application` COM 可用；WPS（KWPS）不可用。
- TrackRevisions+TrackFormatting 下，字体/字号/对齐/缩进改动全部产生修订，导出 PDF 可见"设置了格式: …"气泡。
- lead_prefix 引题拆分需**先套整段格式、再套前缀格式**（顺序反了前缀会被覆盖——revise_format.py 已修正）。
- **`Font.NameFarEast` 不能赋 "Times New Roman" 等非东亚字体名**（COM 报 0x800a16d4）——纯英文段落的
  EN 元素规则不设 east_asian，拉丁渲染由 Font.Name（ascii）负责。
- 表格/段落映射键同时接受 "1" 与 "t1"/"p1" 两种写法（脚本已兼容）。
- COM 里 Bold 读值为 -1(True)/0(False)/9999999(MIXED)；文档已有作者未处理文字修订时，Range.Text 会同时
  含插入与删除文本，格式修订可正常叠加，切勿 Accept/Reject。
- 实战案例：SDR 英文期刊稿（824段/4表/34图）——英文标题按 GB/T EN 元素表校准（题名14pt加粗居中、作者10.5、
  摘要/关键词引题加粗+内容9pt、章标题12pt加粗、图表题注9pt加粗居中、注释9pt两端对齐、参考文献9pt），
  226段+4表映射，约480条格式修订零失败；英文稿"黑体"层级须转译为加粗（英文无黑体）。
- 测试样例（中文）：15 段制冷主题论文，p1→黑体18pt、摘要引题黑体9pt/正文仿宋9pt、图题居中黑体9pt、
  章标题顶格左对齐，全部通过复核。
