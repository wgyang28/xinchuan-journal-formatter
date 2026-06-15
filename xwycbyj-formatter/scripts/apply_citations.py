#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
apply_citations.py — 依据 citation_map.py 把引文写回（脚注改写 / 括注删除或转脚注）。

输入：apply_format.py 产出的 *_新传研究投稿版.docx（已完成机械格式）。
本脚本就地写回该文件，返回变更数据（内存），不产生任何 JSON 文件。

铁律：只搬运既有信息；新建/改写文本均来自 citation_map.py 的人工判断表，不在此处编造。
"""

import os
import sys
import json
import argparse

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from footnote_ops import FootnoteDoc, w  # noqa: E402
import citation_map as M  # noqa: E402


def _footnote_spacing_cfg():
    """从 references/format_config.json 读脚注段前/段后/行距。"""
    cfg_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                            "references", "format_config.json")
    try:
        fn = json.load(open(cfg_path, encoding="utf-8"))["footnote"]
        return (fn.get("space_before_pt", 0), fn.get("space_after_pt", 0),
                str(fn.get("line_spacing", "single")).lower() == "single")
    except Exception:
        return 0, 0, True


def run(docx):
    """就地写回引文，返回变更日志 dict（不产 JSON）。"""
    fd = FootnoteDoc(docx)
    before = fd.integrity()
    paras = list(fd.doc.iter(w("p")))
    log = {"reformatted": [], "deleted_fn": [], "intext": [], "warnings": []}

    def para_text(p):
        return "".join(t.text or "" for t in p.iter(w("t")))

    def get_para(idx, cite):
        """优先按索引取段；若该段不含 cite，则全文搜索并告警。"""
        if 0 <= idx < len(paras) and cite in para_text(paras[idx]):
            return paras[idx]
        for p in paras:
            if cite in para_text(p):
                log["warnings"].append("段索引 P%d 未含 %s，已改用全文匹配定位。" % (idx, cite))
                return p
        return None

    # ── 1) 改写现有脚注 ──────────────────────────────────────────────────────
    for fn_id, segs in M.FOOTNOTE_TARGETS.items():
        try:
            fd.set_footnote_runs(fn_id, segs)
            log["reformatted"].append({"fn": fn_id, "text": "".join(s for s, _ in segs)})
        except KeyError:
            log["warnings"].append("脚注 fn%s 不存在，跳过改写。" % fn_id)

    # ── 2) 删除空脚注 ────────────────────────────────────────────────────────
    for fn_id in M.DELETE_FN:
        fd.delete_footnote(fn_id)
        log["deleted_fn"].append(fn_id)

    # ── 3) 括注动作 ──────────────────────────────────────────────────────────
    for act in M.INTEXT_ACTIONS:
        idx, cite, action = act["para"], act["cite"], act["action"]
        if action == "keep_flag":
            log["intext"].append({"para": idx, "cite": cite, "action": "保留待复核"})
            continue
        if action == "keep":
            log["intext"].append({"para": idx, "cite": cite, "action": "保留括注"})
            continue
        p = get_para(idx, cite)
        if p is None:
            log["warnings"].append("找不到括注：P%d %s" % (idx, cite))
            continue
        if action == "delete":
            ok = fd.delete_text_in_para(p, cite)
            log["intext"].append({"para": idx, "cite": cite, "action": "删除括注" if ok else "删除失败"})
        elif action == "to_footnote":
            if act.get("fn_segments"):
                new_id = fd.create_footnote("·")
                fd.set_footnote_runs(new_id, act["fn_segments"])
                ftext = "".join(s for s, _ in act["fn_segments"])
            else:
                ftext = act["fn_text"]
                new_id = fd.create_footnote(ftext)
            ins = fd.insert_fnref_after_anchor(p, cite, new_id)
            dele = fd.delete_text_in_para(p, cite)
            log["intext"].append({"para": idx, "cite": cite, "action": "转脚注",
                                  "new_fn": new_id, "fn_text": ftext,
                                  "ref_inserted": ins, "cite_deleted": dele})

    # ── 4) 脚注序号位置统一到句末标点之前 ─────────────────────────────────────
    moved = fd.fix_fnref_punctuation()
    log["fnref_punct_moved"] = moved

    # ── 4.5) 脚注段前/段后/行距（直接写 footnotes.xml，确保 before 落属性） ─────
    b, a, single = _footnote_spacing_cfg()
    fd.set_footnote_para_spacing(b, a, single)
    log["footnote_spacing"] = {"before_pt": b, "after_pt": a, "single": single}

    # ── 5) 校验 + 保存 ───────────────────────────────────────────────────────
    after = fd.integrity()
    new_missing = sorted(set(after["missing"]) - set(before["missing"]))
    new_orphan = sorted(set(after["orphan"]) - set(before["orphan"]))
    fd.save()

    log["integrity"] = {
        "before": before, "after": after,
        "new_missing": new_missing, "new_orphan": new_orphan,
        "ok": (not new_missing and not new_orphan and not after["gaps"]),
    }
    return log


def print_summary(log):
    after = log["integrity"]["after"]
    before = log["integrity"]["before"]
    print("改写脚注 %d 条 | 删除空脚注 %s | 括注处理 %d 条"
          % (len(log["reformatted"]), log["deleted_fn"], len(log["intext"])))
    n_del = sum(1 for x in log["intext"] if x["action"] == "删除括注")
    n_to = sum(1 for x in log["intext"] if x["action"] == "转脚注")
    n_keep = sum(1 for x in log["intext"] if x["action"] == "保留待复核")
    n_keep_note = sum(1 for x in log["intext"] if x["action"] == "保留括注")
    print("  括注删除 %d | 转脚注 %d | 保留括注 %d | 保留待复核 %d" % (n_del, n_to, n_keep_note, n_keep))
    print("脚注序号移到句末标点之前：%d 处" % log.get("fnref_punct_moved", 0))
    print("完整性：脚注 %d→%d，正文引用 %d→%d，断号=%s"
          % (before["footnote_count"], after["footnote_count"],
             before["body_ref_count"], after["body_ref_count"], after["gaps"]))
    print("新增缺失=%s 新增孤立=%s → %s"
          % (log["integrity"]["new_missing"], log["integrity"]["new_orphan"],
             "完整性 PASS ✅" if log["integrity"]["ok"] else "需检查 ❌"))
    for wn in log["warnings"]:
        print("  ⚠ " + wn)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx", help="apply_format 产出的 *_新传研究投稿版.docx")
    args = ap.parse_args()
    if not os.path.exists(args.docx):
        print("文件不存在: %s" % args.docx, file=sys.stderr)
        sys.exit(1)
    print_summary(run(args.docx))


if __name__ == "__main__":
    main()
