# -*- coding: utf-8 -*-
"""Word 文档格式体检（只读，经 Word COM）。

输出逐段格式清单 JSON，供 LLM 对照 GB/T 7713.2—2022 规范做元素分类，
再生成 revise_format.py 所需的 mapping.json。

用法:
    python inspect_docx.py <docx> [--out result.json]

要点:
- 段号 p<N> 为 Word Paragraphs 集合的 1-based 序号（含表格内段落），
  与 revise_format.py 使用同一枚举，可直接复用。
- 需要本机安装 Word；文件若正被 Word/WPS 打开会失败，请先关闭。
- 属性值为 "MIXED" 表示该段内格式不一致（Word 返回 wdUndefined=9999999）。
"""
import argparse
import json
import os
import sys

WD_UNDEFINED = 9999999
WD_WITH_IN_TABLE = 12
ALIGN = {0: "left", 1: "center", 2: "right", 3: "justify", 4: "distribute"}


def clean_text(t):
    return t.replace("\r", "⏎").replace("\x07", "¶").replace("\x0b", "⏎")


def fmt_val(v, transform=None):
    if v is None:
        return None
    if isinstance(v, float) and v == int(v):
        v = int(v)
    if v == WD_UNDEFINED:
        return "MIXED"
    return transform(v) if transform else v


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("--out", default=None, help="结果写入该 JSON 文件（默认打印 stdout）")
    args = ap.parse_args()

    path = os.path.abspath(args.docx)
    if not os.path.isfile(path):
        print(json.dumps({"error": f"file not found: {path}"}))
        sys.exit(2)

    import pythoncom
    import win32com.client as win32

    pythoncom.CoInitialize()
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    result = {}
    try:
        try:
            doc = word.Documents.Open(path, ReadOnly=True, AddToRecentFiles=False)
        except Exception as e:
            print(json.dumps({"error": f"open failed: {e}",
                              "hint": "文件可能正被 Word/WPS 打开，请先关闭后重试"}))
            sys.exit(3)
        try:
            ps = doc.Sections(1).PageSetup
            result["page"] = {
                "width_mm": round(ps.PageWidth / 72 * 25.4, 1),
                "height_mm": round(ps.PageHeight / 72 * 25.4, 1),
                "is_A4": abs(ps.PageWidth - 595.3) < 3 and abs(ps.PageHeight - 841.9) < 3,
            }
            result["counts"] = {
                "paragraphs": doc.Paragraphs.Count,
                "tables": doc.Tables.Count,
                "inline_shapes": doc.InlineShapes.Count,
            }
            result["track_revisions_on"] = bool(doc.TrackRevisions)
            paras = []
            for i in range(1, doc.Paragraphs.Count + 1):
                para = doc.Paragraphs.Item(i)
                rng = para.Range
                f = rng.Font
                pf = para.Format
                try:
                    style = para.Style.NameLocal
                except Exception:
                    style = "?"
                paras.append({
                    "p": i,
                    "in_table": bool(rng.Information(WD_WITH_IN_TABLE)),
                    "text": clean_text(rng.Text)[:80],
                    "style": style,
                    "align": fmt_val(pf.Alignment, lambda v: ALIGN.get(v, v)),
                    "first_line_indent_chars": fmt_val(pf.CharacterUnitFirstLineIndent),
                    "ea_font": fmt_val(f.NameFarEast),
                    "ascii_font": fmt_val(f.Name),
                    "size_pt": fmt_val(f.Size),
                    "bold": fmt_val(f.Bold),
                })
            result["paragraphs"] = paras
        finally:
            doc.Close(SaveChanges=0)
    finally:
        word.Quit()
        pythoncom.CoUninitialize()

    out = json.dumps(result, ensure_ascii=False, indent=1)
    if args.out:
        with open(args.out, "w", encoding="utf-8") as fh:
            fh.write(out)
        print(f"written {args.out}; paragraphs={len(result.get('paragraphs', []))}")
    else:
        print(out)


if __name__ == "__main__":
    main()
