#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
footnote_ops.py — 稳健的 docx 脚注 XML 操作库。

为什么不用 python-docx：python-docx 不能干净地新建脚注。本库直接操作
word/footnotes.xml + word/document.xml，并以「克隆现有脚注/引用节点作模板」
的方式新建，避免硬编码随文档而变的样式 id（如 ae / af0），跨文档稳健。

提供能力（供引文转换脚本 import 调用）：
  - get_footnote_text / set_footnote_text   读取、整体替换脚注文本（保留编号标记）
  - create_footnote(text) -> id             克隆模板新建脚注，返回新 id
  - insert_fnref_after_anchor(p, anchor, id) 在正文锚点后插入脚注引用标记
  - delete_text_in_para(p, target)          删除正文括注文本（只清 w:t，不删 run）
  - integrity()                             校验 正文引用↔脚注 一致、ID 连续
  - save(out)                               写回（脚注按 id 排序，WPS 兼容）

铁律：本库只搬运/删除既有文本，不生成、不补全任何文献内容。

命令行自测：python3 footnote_ops.py selftest <some.docx>
"""

import os
import sys
import copy
import zipfile
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def w(tag):
    return f"{{{W}}}{tag}"


class FootnoteDoc:
    def __init__(self, path):
        self.path = path
        with zipfile.ZipFile(path) as z:
            self.files = {n: z.read(n) for n in z.namelist()}
        self.doc = etree.fromstring(self.files["word/document.xml"])
        self.fns = (etree.fromstring(self.files["word/footnotes.xml"])
                    if "word/footnotes.xml" in self.files else None)
        self.settings = (etree.fromstring(self.files["word/settings.xml"])
                         if "word/settings.xml" in self.files else None)
        if self.fns is None:
            raise ValueError("文档不含 word/footnotes.xml；新建脚注需文档已有脚注基础结构。")

    # ── 脚注集合 ────────────────────────────────────────────────────────────
    def _all_fn(self):
        return self.fns.findall(w("footnote"))

    def _real_fn(self):
        return [fn for fn in self._all_fn() if fn.get(w("id")) not in ("-1", "0")]

    def max_id(self):
        ids = [int(fn.get(w("id"))) for fn in self._all_fn()
               if (fn.get(w("id")) or "").lstrip("-").isdigit()]
        return max(ids) if ids else 0

    def _find_fn(self, fn_id):
        for fn in self._all_fn():
            if fn.get(w("id")) == str(fn_id):
                return fn
        return None

    @staticmethod
    def _text_runs(fn):
        """脚注内除『编号标记 run』外的文本 run。"""
        out = []
        for r in fn.iter(w("r")):
            if r.find(w("footnoteRef")) is not None:
                continue
            out.append(r)
        return out

    def get_footnote_text(self, fn_id):
        fn = self._find_fn(fn_id)
        if fn is None:
            return None
        return "".join(t.text or "" for t in fn.iter(w("t")))

    def set_footnote_text(self, fn_id, new_text):
        """整体替换脚注文本，保留编号标记 run。"""
        fn = self._find_fn(fn_id)
        if fn is None:
            raise KeyError(f"脚注 {fn_id} 不存在")
        text_runs = self._text_runs(fn)
        # 清空所有文本 run
        for r in text_runs:
            for t in r.findall(w("t")):
                t.text = ""
        if text_runs:
            r0 = text_runs[0]
            t = r0.find(w("t"))
            if t is None:
                t = etree.SubElement(r0, w("t"))
            t.set(XML_SPACE, "preserve")
            t.text = " " + new_text  # 与编号间留一空格
        else:
            # 无文本 run：在编号 run 后补一个
            p = fn.find(w("p"))
            r = etree.SubElement(p, w("r"))
            t = etree.SubElement(r, w("t"))
            t.set(XML_SPACE, "preserve")
            t.text = " " + new_text
        return True

    def set_footnote_runs(self, fn_id, segments):
        """按 segments=[(text, italic_bool), ...] 分段写脚注文本，保留编号标记。
        用于外文文献：题名段 italic=True 即可写成斜体。"""
        fn = self._find_fn(fn_id)
        if fn is None:
            raise KeyError(f"脚注 {fn_id} 不存在")
        p = fn.find(w("p"))
        text_runs = self._text_runs(fn)
        template_rpr = None
        for r in text_runs:
            rpr = r.find(w("rPr"))
            if rpr is not None:
                template_rpr = copy.deepcopy(rpr)
                break
        ref_run = None
        for r in p.findall(w("r")):
            if r.find(w("footnoteRef")) is not None:
                ref_run = r
                break
        for r in text_runs:
            r.getparent().remove(r)
        anchor = ref_run
        first = True
        for text, italic in segments:
            nr = etree.Element(w("r"))
            rpr = copy.deepcopy(template_rpr) if template_rpr is not None else etree.Element(w("rPr"))
            i_el = rpr.find(w("i"))
            if italic and i_el is None:
                etree.SubElement(rpr, w("i"))
            elif not italic and i_el is not None:
                rpr.remove(i_el)
            nr.append(rpr)
            t = etree.SubElement(nr, w("t"))
            t.set(XML_SPACE, "preserve")
            t.text = (" " + text) if first else text
            first = False
            if anchor is not None:
                anchor.addnext(nr)
            else:
                p.append(nr)
            anchor = nr
        return True

    def delete_footnote(self, fn_id):
        """删除一条脚注节点，移除正文引用标记 run，并清除 settings.xml 中对该 id 的引用
        （避免悬空引用触发 Word“不可读取的内容”）。"""
        fn = self._find_fn(fn_id)
        if fn is not None:
            self.fns.remove(fn)
        for ref in list(self.doc.iter(w("footnoteReference"))):
            if ref.get(w("id")) == str(fn_id):
                run = ref.getparent()
                if run is not None and run.getparent() is not None:
                    run.getparent().remove(run)
        # settings.xml: w:footnotePr/w:footnote[@w:id] 分隔符引用
        if self.settings is not None:
            fnpr = self.settings.find(w("footnotePr"))
            if fnpr is not None:
                for ref in fnpr.findall(w("footnote")):
                    if ref.get(w("id")) == str(fn_id):
                        fnpr.remove(ref)
        return True

    def create_footnote(self, text):
        """克隆一个现有真实脚注作模板，新建脚注；返回新 id。"""
        template = None
        for fn in self._real_fn():
            if self._text_runs(fn):
                template = fn
                break
        if template is None:
            raise ValueError("找不到可用作模板的现有脚注。")
        new = copy.deepcopy(template)
        new_id = self.max_id() + 1
        new.set(w("id"), str(new_id))
        # 用 set_footnote_text 同样的逻辑写入文本
        for r in self._text_runs(new):
            for t in r.findall(w("t")):
                t.text = ""
        tr = self._text_runs(new)
        t = tr[0].find(w("t"))
        if t is None:
            t = etree.SubElement(tr[0], w("t"))
        t.set(XML_SPACE, "preserve")
        t.text = " " + text
        self.fns.append(new)
        return new_id

    # ── 正文引用标记 ──────────────────────────────────────────────────────────
    def _fnref_template(self):
        """从正文里找一个现有 footnoteReference 所在 run 作模板。"""
        for r in self.doc.iter(w("r")):
            if r.find(w("footnoteReference")) is not None:
                return r
        return None

    def insert_fnref_after_anchor(self, para, anchor, fn_id, occurrence=1):
        """在段落 para 中 anchor 文本结束处之后，插入指向 fn_id 的脚注引用 run。"""
        tmpl = self._fnref_template()
        if tmpl is None:
            raise ValueError("文档中找不到现有脚注引用标记作模板。")
        # 构造 w:t 节点的 (节点, 在拼接串中的起点) 映射
        tmap = []
        pos = 0
        for t in para.iter(w("t")):
            s = t.text or ""
            tmap.append((t, pos, pos + len(s)))
            pos += len(s)
        full = "".join(t.text or "" for t, _, _ in tmap)
        idx = -1
        for _ in range(occurrence):
            idx = full.find(anchor, idx + 1)
            if idx == -1:
                return False
        end = idx + len(anchor)
        # 找到包含 end 位置的 w:t 节点，定位其所属 run
        target_t = None
        for t, a, b in tmap:
            if a < end <= b:
                target_t = t
                break
        if target_t is None and tmap:
            target_t = tmap[-1][0]
        # 该 w:t 的祖先 run
        run = target_t
        while run is not None and run.tag != w("r"):
            run = run.getparent()
        new_run = copy.deepcopy(tmpl)
        ref = new_run.find(w("footnoteReference"))
        ref.set(w("id"), str(fn_id))
        run.addnext(new_run)
        return True

    # ── 删除正文括注文本（只清 w:t，不删 run，保住内嵌脚注标记） ───────────────
    def delete_text_in_para(self, para, target, occurrence=1):
        tmap = []
        pos = 0
        for t in para.iter(w("t")):
            s = t.text or ""
            tmap.append([t, pos, pos + len(s)])
            pos += len(s)
        full = "".join(t.text or "" for t, _, _ in tmap)
        idx = -1
        for _ in range(occurrence):
            idx = full.find(target, idx + 1)
            if idx == -1:
                return False
        start, end = idx, idx + len(target)
        for t, a, b in tmap:
            if b <= start or a >= end:
                continue
            s = t.text or ""
            cut_lo = max(start, a) - a
            cut_hi = min(end, b) - a
            t.text = s[:cut_lo] + s[cut_hi:]
        return True

    # ── 脚注序号位置：句末标点之前 ────────────────────────────────────────────
    def fix_fnref_punctuation(self, move_set="。！？，；："):
        """把位于句末标点【之后】的脚注引用标记移到标点【之前】。
        规范：注释序号统一置于包含引文的句子/词之后、标点之前（……研究①。而非……研究。①）。
        仅当引用标记紧跟在 move_set 中的标点之后才移动；引号、括号、书名号不动。
        返回移动次数。"""
        moved = 0
        for p in self.doc.iter(w("p")):
            runs = p.findall(w("r"))
            for i, r in enumerate(runs):
                if r.find(w("footnoteReference")) is None:
                    continue
                prev = None
                for j in range(i - 1, -1, -1):
                    tnode = runs[j].find(w("t"))
                    if tnode is not None and (tnode.text or ""):
                        prev = runs[j]
                        break
                if prev is None:
                    continue
                t = prev.find(w("t"))
                if not t.text or t.text[-1] not in move_set:
                    continue
                punct = t.text[-1]
                t.text = t.text[:-1]
                newr = etree.Element(w("r"))
                rpr = prev.find(w("rPr"))
                if rpr is not None:
                    newr.append(copy.deepcopy(rpr))
                nt = etree.SubElement(newr, w("t"))
                nt.set(XML_SPACE, "preserve")
                nt.text = punct
                r.addnext(newr)
                moved += 1
        return moved

    # ── 脚注段落间距/行距 ─────────────────────────────────────────────────────
    def set_footnote_para_spacing(self, before_pt=0, after_pt=0, single=True):
        """直接在 footnotes.xml 的每条脚注段落 w:pPr/w:spacing 上设置段前/段后与行距。
        before/after 以磅(pt)计（×20 转 twips）；single=True 设单倍行距(line=240,auto)。
        一并清零 beforeLines/afterLines 与自动间距。分隔符脚注(-1/0)不动。"""
        bt = str(int(round(before_pt * 20)))
        at = str(int(round(after_pt * 20)))
        for fn in self._real_fn():
            for p in fn.findall(w("p")):
                pPr = p.find(w("pPr"))
                if pPr is None:
                    pPr = etree.Element(w("pPr"))
                    p.insert(0, pPr)
                spacing = pPr.find(w("spacing"))
                if spacing is None:
                    spacing = etree.Element(w("spacing"))
                    pstyle = pPr.find(w("pStyle"))
                    if pstyle is not None:
                        pstyle.addnext(spacing)   # pStyle 之后，符合 pPr schema 顺序
                    else:
                        pPr.insert(0, spacing)
                spacing.set(w("before"), bt)
                spacing.set(w("after"), at)
                for a in ("beforeLines", "afterLines", "beforeAutospacing", "afterAutospacing"):
                    spacing.set(w(a), "0")
                if single:
                    spacing.set(w("line"), "240")
                    spacing.set(w("lineRule"), "auto")
        return True

    # ── 校验 ─────────────────────────────────────────────────────────────────
    def integrity(self):
        body_refs = {r.get(w("id")) for r in self.doc.iter(w("footnoteReference"))}
        fn_ids = {fn.get(w("id")) for fn in self._real_fn()}
        missing = sorted(body_refs - fn_ids, key=lambda x: int(x) if x.lstrip("-").isdigit() else 0)
        orphan = sorted(fn_ids - body_refs, key=lambda x: int(x) if x.lstrip("-").isdigit() else 0)
        ints = sorted(int(i) for i in fn_ids if i.lstrip("-").isdigit())
        gaps = [i for i in range(ints[0], ints[-1] + 1) if i not in set(ints)] if ints else []
        return {"missing": missing, "orphan": orphan, "gaps": gaps,
                "body_ref_count": len(body_refs), "footnote_count": len(fn_ids)}

    # ── 写回 ─────────────────────────────────────────────────────────────────
    def save(self, out=None):
        out = out or self.path
        # 脚注按 id 排序（WPS 兼容）
        all_fn = self._all_fn()
        for fn in all_fn:
            self.fns.remove(fn)
        for fn in sorted(all_fn, key=lambda x: int(x.get(w("id"))) if (x.get(w("id")) or "").lstrip("-").isdigit() else 0):
            self.fns.append(fn)
        self.files["word/document.xml"] = etree.tostring(
            self.doc, xml_declaration=True, encoding="UTF-8", standalone=True)
        self.files["word/footnotes.xml"] = etree.tostring(
            self.fns, xml_declaration=True, encoding="UTF-8", standalone=True)
        if self.settings is not None:
            self.files["word/settings.xml"] = etree.tostring(
                self.settings, xml_declaration=True, encoding="UTF-8", standalone=True)
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
            for n, d in self.files.items():
                z.writestr(n, d)
        return out


# ── 自测 ─────────────────────────────────────────────────────────────────────
def selftest(src):
    import shutil
    tmp = os.path.splitext(src)[0] + "_selftest.docx"
    shutil.copy2(src, tmp)
    print("自测副本：", tmp)

    fd = FootnoteDoc(tmp)
    before = fd.integrity()
    print("初始：脚注 %d，正文引用 %d，max_id=%d，缺失=%s 孤立=%s 断号=%s"
          % (before["footnote_count"], before["body_ref_count"], fd.max_id(),
             before["missing"], before["orphan"], before["gaps"]))

    # 1) 新建脚注
    marker = "【自测】田野笔记，2026年6月2日。"
    new_id = fd.create_footnote(marker)
    print("新建脚注 id=%d" % new_id)

    # 2) 把引用标记插入第一处有文本的正文段落（锚到前若干字）
    inserted = False
    for p in fd.doc.iter(w("p")):
        txt = "".join(t.text or "" for t in p.iter(w("t")))
        if len(txt.strip()) >= 6 and not txt.strip().startswith(("摘要", "关键词", "内容提要")):
            anchor = txt.strip()[:6]
            if fd.insert_fnref_after_anchor(p, anchor, new_id):
                inserted = True
                print("引用标记已插入锚点：%r" % anchor)
                break
    if not inserted:
        print("✗ 未能插入引用标记")

    fd.save()

    # 3) 往返复读。只追究『本次操作新引入』的缺失/孤立；既存孤立（如空脚注 fn1）不计入。
    fd2 = FootnoteDoc(tmp)
    after = fd2.integrity()
    rt_text = fd2.get_footnote_text(new_id)
    new_missing = set(after["missing"]) - set(before["missing"])
    new_orphan = set(after["orphan"]) - set(before["orphan"])
    ok = (not new_missing and not new_orphan
          and after["footnote_count"] == before["footnote_count"] + 1
          and rt_text and marker in rt_text)
    if before["orphan"]:
        print("注：原稿已存在孤立脚注（无正文引用）：%s —— 多为空脚注，建议清理。" % before["orphan"])
    print("往返后：脚注 %d，引用 %d，缺失=%s 孤立=%s 断号=%s"
          % (after["footnote_count"], after["body_ref_count"],
             after["missing"], after["orphan"], after["gaps"]))
    print("复读新脚注文本：%r" % (rt_text,))
    print("\n自测结果：", "PASS ✅" if ok else "FAIL ❌")
    os.remove(tmp)
    return ok


if __name__ == "__main__":
    if len(sys.argv) >= 3 and sys.argv[1] == "selftest":
        sys.exit(0 if selftest(sys.argv[2]) else 1)
    print(__doc__)
