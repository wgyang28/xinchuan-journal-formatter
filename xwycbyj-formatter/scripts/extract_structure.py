#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
extract_structure.py —《新闻与传播研究》格式 skill 的「结构 + 引文」抽取器。

职责（混合策略里的「代码可靠抽取」环节）：
  - 只读不写。把 .docx 的结构与各类引文线索抽成 JSON，交给 Claude 做格式映射判断。
  - 绝不修改文档、绝不补全/推断任何文献内容。

抽取内容：
  1. 正文段落与标题层级（按 format_config.json 的前缀正则判定）
  2. 前置部分线索：标题区、内容提要、关键词、英文标题/摘要
  3. 全部脚注文本（footnotes.xml）
  4. 正文括注（夹注）——含年份的括注，并标记同段是否已有脚注引用
  5. 文末参考文献列表（若存在）

用法：
  python3 extract_structure.py <输入.docx> [-o 输出.json]
  不带 -o 时，打印人读摘要并把 JSON 写到 <输入名>.structure.json
"""

import sys
import os
import re
import json
import zipfile
import argparse
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def w(tag):
    return f"{{{W}}}{tag}"


# ── 配置加载（标题前缀正则从 references/format_config.json 读取，不写死） ──────────
def load_heading_patterns():
    cfg_path = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        "references", "format_config.json",
    )
    default = {
        "h1": r"^[一二三四五六七八九十百]+、",
        "h2": r"^（[一二三四五六七八九十]+）",
        "h3": r"^\d+[．.]",
        "h4": r"^（\d+）",
    }
    try:
        with open(cfg_path, encoding="utf-8") as f:
            cfg = json.load(f)
        out = {}
        for lvl, spec in cfg.get("headings", {}).items():
            out[lvl] = spec.get("prefix_regex", default.get(lvl, ""))
        return out or default
    except Exception:
        return default


HEADING_PATTERNS = load_heading_patterns()


CJK = re.compile(r"[一-鿿]")


def heading_level(text):
    t = text.strip()
    # 标题须含汉字（排除表格小数 0.796、纯英文公式等）且长度受限
    if not CJK.search(t) or len(t) > 40:
        return 0
    for lvl in ("h1", "h2", "h3", "h4"):
        pat = HEADING_PATTERNS.get(lvl)
        if pat and re.match(pat, t):
            return int(lvl[1])
    return 0


# ── docx 读取 ────────────────────────────────────────────────────────────────
def read_xml(z, name):
    if name in z.namelist():
        return etree.fromstring(z.read(name))
    return None


def para_text(p):
    return "".join(t.text or "" for t in p.iter(w("t")))


def para_footnote_refs(p):
    return [r.get(w("id")) for r in p.iter(w("footnoteReference"))]


def para_style(p):
    ppr = p.find(w("pPr"))
    if ppr is None:
        return None
    pstyle = ppr.find(w("pStyle"))
    return pstyle.get(w("val")) if pstyle is not None else None


# ── 引文线索：括注（夹注） ────────────────────────────────────────────────────
# 含 4 位年份的全角括注（中文夹注 / APA 中译）
CN_YEAR_CITE = re.compile(r"（[^（）]{0,100}?\d{4}[a-z]?[^（）]{0,100}?）")
# 英文夹注 (Author, 2020) / (Author & Author, 2020) / (Author et al., 2020)
EN_YEAR_CITE = re.compile(r"\([^()]{0,120}?\b(1[5-9]\d{2}|20\d{2})[a-z]?\b[^()]{0,60}?\)")
# 田野/访谈类括注：（YYYYMMDD，……）或（……，田野笔记/访谈……）
FIELD_CITE = re.compile(r"（[^（）]{0,40}?(田野|访谈|笔记|聊天记录|观察)[^（）]{0,40}?）")
# 纯外文人名括注（应保留，不算引文）：（Halbwachs）、（Fentress & Wickham）
NAME_ONLY = re.compile(r"^（[A-Za-z][A-Za-z.&,\-\s]{1,40}）$")


def classify_citation(s):
    if FIELD_CITE.search(s):
        return "field_or_interview"
    if NAME_ONLY.match(s):
        return "foreign_name_keep"
    if re.search(r"[A-Za-z]", s) and not re.search(r"[一-鿿]", s):
        return "english_intext"
    return "chinese_intext"


def find_intext_citations(body):
    hits = []
    for i, p in enumerate(body.iter(w("p"))):
        text = para_text(p)
        if not text:
            continue
        has_fn = bool(para_footnote_refs(p))
        seen = set()
        for rx in (CN_YEAR_CITE, EN_YEAR_CITE, FIELD_CITE):
            for m in rx.finditer(text):
                s = m.group()
                if s in seen:
                    continue
                seen.add(s)
                start = max(0, m.start() - 25)
                end = min(len(text), m.end() + 25)
                hits.append({
                    "para_idx": i,
                    "match": s,
                    "kind_guess": classify_citation(s),
                    "same_para_has_footnote": has_fn,
                    "context": text[start:end],
                })
    return hits


# ── 引文线索：文末参考文献列表 ────────────────────────────────────────────────
REF_HEADING = re.compile(r"^\s*(参考文献|引用文献|参考书目|征引文献|References|REFERENCES|Bibliography)\s*[:：]?\s*$")


def find_end_references(paras_text):
    """paras_text: list[(idx, text)]。返回 {heading_idx, entries:[...]} 或 None。"""
    head_idx = None
    for idx, text in paras_text:
        if REF_HEADING.match(text.strip()):
            head_idx = idx
    if head_idx is None:
        return None
    entries = []
    for idx, text in paras_text:
        if idx <= head_idx:
            continue
        t = text.strip()
        if not t:
            continue
        # 资助/作者信息等收尾段落不计入；只收形似文献条目的段落
        entries.append({"idx": idx, "text": t})
    return {"heading_idx": head_idx, "entries": entries}


# ── 前置部分线索 ──────────────────────────────────────────────────────────────
def find_front_matter(paras_text):
    fm = {
        "abstract_idx": None, "abstract_label": None,
        "keywords_idx": None, "keywords_label": None,
        "has_english_title": False, "has_english_abstract": False,
        "english_abstract_idx": None,
        "title_candidates": [],
    }
    for idx, text in paras_text[:8]:
        t = text.strip()
        if t and idx < 6 and not re.match(r"^(内容提要|摘要|关键词|关\s*键\s*词)", t):
            fm["title_candidates"].append({"idx": idx, "text": t})
    for idx, text in paras_text[:40]:
        t = text.strip()
        m = re.match(r"^(内容提要|摘\s*要|提\s*要)\s*[:：]?", t)
        if m and fm["abstract_idx"] is None:
            fm["abstract_idx"] = idx
            fm["abstract_label"] = m.group(1).replace(" ", "")
        m = re.match(r"^(关\s*键\s*词|關鍵詞)\s*[:：]?", t)
        if m and fm["keywords_idx"] is None:
            fm["keywords_idx"] = idx
            fm["keywords_label"] = m.group(1)
        # 英文摘要：连续英文且较长
        if (re.match(r"^(Abstract|ABSTRACT)\b", t) or
                (len(t) > 120 and len(re.findall(r"[A-Za-z]", t)) > 0.7 * len(t))):
            if fm["english_abstract_idx"] is None:
                fm["has_english_abstract"] = True
                fm["english_abstract_idx"] = idx
        if re.match(r"^[A-Z][A-Za-z].*[A-Za-z]$", t) and len(t.split()) >= 4 and idx < 20:
            fm["has_english_title"] = True
    return fm


# ── 主流程 ───────────────────────────────────────────────────────────────────
def extract(path):
    with zipfile.ZipFile(path) as z:
        doc = read_xml(z, "word/document.xml")
        fns = read_xml(z, "word/footnotes.xml")
        ens = read_xml(z, "word/endnotes.xml")

    body = doc.find(w("body")) if doc is not None else None
    if body is None:
        raise ValueError("无法解析 word/document.xml")

    paras = list(body.iter(w("p")))
    paras_text = [(i, para_text(p)) for i, p in enumerate(paras)]

    # 表格内段落不参与标题判定（表格小数会误命中 h3）
    table_pids = {id(tp) for tbl in body.iter(w("tbl")) for tp in tbl.iter(w("p"))}

    headings = []
    for i, p in enumerate(paras):
        if id(p) in table_pids:
            continue
        text = para_text(p)
        lvl = heading_level(text)
        if lvl:
            headings.append({"idx": i, "level": lvl, "text": text.strip(), "style": para_style(p)})

    footnotes = []
    if fns is not None:
        for fn in fns.findall(w("footnote")):
            fid = fn.get(w("id"))
            if fid in ("-1", "0"):
                continue
            text = "".join(t.text or "" for t in fn.iter(w("t")))
            footnotes.append({"id": fid, "text": text.strip()})

    endnotes = []
    if ens is not None:
        for en in ens.findall(w("endnote")):
            eid = en.get(w("id"))
            if eid in ("-1", "0"):
                continue
            text = "".join(t.text or "" for t in en.iter(w("t")))
            endnotes.append({"id": eid, "text": text.strip()})

    intext = find_intext_citations(body)
    end_refs = find_end_references(paras_text)
    front = find_front_matter(paras_text)

    # 引文形式判定（供 SKILL 流程分支用）
    modes = []
    if footnotes:
        modes.append("footnote")
    if [h for h in intext if h["kind_guess"] in ("chinese_intext", "english_intext")]:
        modes.append("intext")
    if end_refs:
        modes.append("endref")
    if endnotes:
        modes.append("endnote")

    return {
        "input": os.path.abspath(path),
        "counts": {
            "paragraphs": len(paras),
            "headings": len(headings),
            "footnotes": len(footnotes),
            "endnotes": len(endnotes),
            "intext_citations": len(intext),
            "end_reference_entries": len(end_refs["entries"]) if end_refs else 0,
        },
        "citation_modes": modes,
        "front_matter": front,
        "headings": headings,
        "footnotes": footnotes,
        "endnotes": endnotes,
        "intext_citations": intext,
        "end_references": end_refs,
    }


def print_summary(data):
    c = data["counts"]
    print("=" * 56)
    print("结构抽取摘要：", os.path.basename(data["input"]))
    print("=" * 56)
    print(f"段落 {c['paragraphs']} | 标题 {c['headings']} | 脚注 {c['footnotes']} "
          f"| 尾注 {c['endnotes']} | 括注 {c['intext_citations']} | 文末参考文献 {c['end_reference_entries']}")
    print(f"引文形式判定：{', '.join(data['citation_modes']) or '未检出'}")
    fm = data["front_matter"]
    print(f"内容提要：{'第%s段' % fm['abstract_idx'] if fm['abstract_idx'] is not None else '未检出'}"
          f"（标签 {fm['abstract_label']}）"
          f" | 关键词：{'第%s段' % fm['keywords_idx'] if fm['keywords_idx'] is not None else '未检出'}")
    print(f"英文标题：{'有' if fm['has_english_title'] else '未检出'}"
          f" | 英文摘要：{'有' if fm['has_english_abstract'] else '未检出'}")
    print("-" * 56)
    print("标题层级：")
    for h in data["headings"]:
        print(f"  [{'　'*(h['level']-1)}H{h['level']}] P{h['idx']}: {h['text']}")
    if data["intext_citations"]:
        print("-" * 56)
        print("括注（夹注）线索：")
        for h in data["intext_citations"]:
            flag = " ⚠已有脚注同段" if h["same_para_has_footnote"] else ""
            print(f"  P{h['para_idx']} [{h['kind_guess']}]{flag}: {h['match']}")
    if data["end_references"]:
        print("-" * 56)
        print(f"文末参考文献（标题在 P{data['end_references']['heading_idx']}，"
              f"共 {len(data['end_references']['entries'])} 条）：")
        for e in data["end_references"]["entries"][:200]:
            print(f"  P{e['idx']}: {e['text']}")
    if data["footnotes"]:
        print("-" * 56)
        print("现有脚注：")
        for fn in data["footnotes"]:
            print(f"  fn{fn['id']}: {fn['text']}")
    print("=" * 56)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("-o", "--out", default=None,
                    help="可选：写出 JSON 路径。默认只打印摘要，不产生文件。")
    args = ap.parse_args()

    if not os.path.exists(args.docx):
        print(f"文件不存在: {args.docx}", file=sys.stderr)
        sys.exit(1)

    data = extract(args.docx)
    print_summary(data)
    # 默认不落盘 JSON；仅当显式 -o 时写出
    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"\nJSON 已写入：{args.out}")


if __name__ == "__main__":
    main()
