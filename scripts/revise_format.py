# -*- coding: utf-8 -*-
"""在 Word 修订模式（Track Changes）下，按 GB/T 7713.2—2022 格式规则应用格式调整。

所有改动以修订（revision）形式记录，作者在 Word 中逐项接受/拒绝，原文不会被静默覆盖。

用法:
    python revise_format.py <docx> --map mapping.json [--rules format_rules.json]
                            [--out out.docx] [--dry-run]

mapping.json 结构（段号 = inspect_docx.py 输出的 p<N>，1-based，含表格内段落）:
{
  "paragraphs": {
    "1": "title_cn",
    "2": {"lead_prefix": "摘要：", "lead": "abstract_cn_head", "element": "abstract_cn_body"},
    "3": "heading_chapter"
  },
  "tables": { "1": "table_body" }
}
- 值为字符串: 整段应用该元素格式。
- 值为对象: lead_prefix 匹配段首文字，前缀应用 lead 元素（引题），其余应用 element 元素。
- tables 键为 Word Tables 集合 1-based 序号，整表应用元素格式。
"""
import argparse
import json
import os
import sys

ALIGN = {"left": 0, "center": 1, "right": 2, "justify": 3}


def _set(getter, attr, value):
    try:
        setattr(getter, attr, value)
        return True
    except Exception as ex:
        return f"跳过({attr}: {ex.__class__.__name__})"


def apply_element(rng, e):
    changed = []
    f = rng.Font
    if e.get("east_asian"):
        r = _set(f, "NameFarEast", e["east_asian"])
        changed.append(f"中文字体={e['east_asian']}" if r is True else str(r))
    if e.get("ascii"):
        r = _set(f, "Name", e["ascii"])
        changed.append(f"西文字体={e['ascii']}" if r is True else str(r))
    if e.get("size_pt"):
        r = _set(f, "Size", e["size_pt"])
        changed.append(f"字号={e['size_pt']}pt" if r is True else str(r))
    b = e.get("bold")
    if b is not None:
        r = _set(f, "Bold", bool(b))
        changed.append(f"加粗={bool(b)}" if r is True else str(r))
    if e.get("align") in ALIGN:
        r = _set(rng.ParagraphFormat, "Alignment", ALIGN[e["align"]])
        changed.append(f"对齐={e['align']}" if r is True else str(r))
    if e.get("clear_first_line_indent"):
        r1 = _set(rng.ParagraphFormat, "CharacterUnitFirstLineIndent", 0)
        _set(rng.ParagraphFormat, "FirstLineIndent", 0)
        changed.append("顶格(清除首行缩进)" if r1 is True else str(r1))
    if e.get("first_line_indent_chars"):
        r = _set(rng.ParagraphFormat, "CharacterUnitFirstLineIndent", e["first_line_indent_chars"])
        changed.append(f"首行缩进{e['first_line_indent_chars']}字符" if r is True else str(r))
    return "; ".join(changed) if changed else "(无格式项)"


def para_idx(key):
    """接受 '12' 或 'p12' 两种段落键，返回 int 或 None。"""
    try:
        return int(key[1:]) if key.startswith("p") else int(key)
    except ValueError:
        return None


def validate_mapping(mapping, elems):
    for key, spec in mapping.get("paragraphs", {}).items():
        if para_idx(key) is None:
            sys.exit(f"mapping 段落键须为数字或 p<数字>: {key}")
        ids = [spec] if isinstance(spec, str) else [spec.get("element"), spec.get("lead")]
        for eid in filter(None, ids):
            if eid not in elems:
                sys.exit(f"未知元素 id: {eid}（可选: {', '.join(elems)}）")
    for key, eid in mapping.get("tables", {}).items():
        try:
            int(key[1:]) if key.startswith("t") else int(key)
        except ValueError:
            sys.exit(f"tables 键须为数字或 t<数字>: {key}")
        if eid not in elems:
            sys.exit(f"未知元素 id: {eid}（可选: {', '.join(elems)}）")


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("--map", required=True, help="mapping.json（段号→元素）")
    ap.add_argument("--rules", default=os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                                    "format_rules.json"))
    ap.add_argument("--out", default=None, help="另存为该文件（默认原文件就地保存，修订可回退）")
    ap.add_argument("--dry-run", action="store_true", help="只打印将执行的改动，不修改文件")
    args = ap.parse_args()

    path = os.path.abspath(args.docx)
    if not os.path.isfile(path):
        sys.exit(f"file not found: {path}")
    if args.out and os.path.exists(os.path.abspath(args.out)) and not args.dry_run:
        sys.exit(f"--out 已存在，拒绝覆盖: {args.out}")

    with open(args.rules, encoding="utf-8") as fh:
        rules = json.load(fh)
    with open(args.map, encoding="utf-8") as fh:
        mapping = json.load(fh)
    elems = {e["id"]: e for e in rules["element_formats"]}
    validate_mapping(mapping, elems)

    import pythoncom
    import win32com.client as win32

    pythoncom.CoInitialize()
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    report = []
    doc = None
    try:
        try:
            doc = word.Documents.Open(path, AddToRecentFiles=False)
        except Exception as e:
            sys.exit(f"打开失败: {e}\n提示: 文件可能正被 Word/WPS 打开，请先关闭。")
        try:
            doc.TrackRevisions = True
            try:
                doc.TrackFormatting = True  # 旧版 Word 无此属性；缺失时格式改动可能不记修订
            except Exception:
                report.append({"warning": "本机 Word 不支持 TrackFormatting，格式改动可能未记为修订"})

            paras = doc.Paragraphs
            total = paras.Count
            for key, spec in mapping.get("paragraphs", {}).items():
                idx = para_idx(key)
                if idx < 1 or idx > total:
                    report.append({"target": key, "error": f"段号越界(共{total}段)"})
                    continue
                para = paras.Item(idx)
                rng = para.Range
                if isinstance(spec, str):
                    eid = spec
                    desc = apply_element(rng, elems[eid])
                else:
                    eid = spec.get("element")
                    desc_parts = []
                    prefix = spec.get("lead_prefix", "")
                    lead_id = spec.get("lead")
                    desc_parts.append(apply_element(rng, elems[eid]))
                    if lead_id and prefix and rng.Text.startswith(prefix):
                        lead_rng = doc.Range(rng.Start, rng.Start + len(prefix))
                        desc_parts.append(f"[{prefix!r}→{lead_id}] " + apply_element(lead_rng, elems[lead_id]))
                    desc = " | ".join(desc_parts)
                report.append({"target": key, "element": eid,
                               "text": rng.Text[:40].strip(), "applied": desc})

            for key, eid in mapping.get("tables", {}).items():
                tidx = int(key[1:]) if key.startswith("t") else int(key)
                if tidx < 1 or tidx > doc.Tables.Count:
                    report.append({"target": key, "error": f"表号越界(共{doc.Tables.Count}表)"})
                    continue
                trng = doc.Tables.Item(tidx).Range
                report.append({"target": key, "element": eid,
                               "applied": apply_element(trng, elems[eid])})

            if args.dry_run:
                print(json.dumps({"dry_run": True, "plan": report}, ensure_ascii=False, indent=1))
            else:
                if args.out:
                    doc.SaveAs2(os.path.abspath(args.out))
                else:
                    doc.Save()
                print(json.dumps({
                    "saved_as": os.path.abspath(args.out) if args.out else path,
                    "track_revisions": bool(doc.TrackRevisions),
                    "revisions_total": doc.Revisions.Count,
                    "applied": report,
                }, ensure_ascii=False, indent=1))
        finally:
            doc.Close(SaveChanges=0)  # 已显式 Save/SaveAs2，此处不重复保存
    finally:
        word.Quit()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
