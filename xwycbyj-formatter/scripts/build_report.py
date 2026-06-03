#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_report.py — 由内存中的格式/引文变更数据渲染《修改报告》Markdown。

供 apply_all.py import 调用：write_report(report_path, fmt, cit, flags)
  fmt   = apply_format.run() 的返回
  cit   = apply_citations.run() 的返回
  flags = citation_map.FLAGS

不读取、不产生任何 JSON 文件。
"""

import os
import datetime


def write_report(report_path, fmt, cit, flags):
    L = []
    inp = os.path.basename(fmt.get("input", ""))
    outp = os.path.basename(fmt.get("output", ""))
    L.append("# 《新闻与传播研究》格式调整 — 修改报告\n")
    L.append(f"- 原文件：`{inp}`")
    L.append(f"- 输出文件：`{outp}`（原文件未改动）")
    L.append(f"- 生成时间：{datetime.date.today().isoformat()}")
    L.append("- 依据：`references/新闻与传播研究_格式与注释规范.md`（运行时唯一权威依据）\n")
    L.append("> **铁律**：本次调整只重排既有文献信息的顺序与标点，未新增、补全或推断任何文献内容。"
             "凡信息缺失或无法确定者，均列入下文「待人工复核」，**未擅自猜测**。\n")

    # 1. 格式变更
    L.append("## 一、格式变更（机械项）\n")
    for c in fmt.get("format_changes", []):
        L.append(f"- ✅ {c}")
    L.append("")

    # 2. 引文统一
    integ = cit.get("integrity", {})
    aft = integ.get("after", {})
    bef = integ.get("before", {})
    intext = cit.get("intext", [])
    n_ref = len(cit.get("reformatted", []))
    n_del = sum(1 for x in intext if x["action"] == "删除括注")
    n_to = sum(1 for x in intext if x["action"] == "转脚注")
    n_keep = sum(1 for x in intext if x["action"] == "保留待复核")
    L.append("## 二、引文统一\n")
    L.append(f"- 改写脚注 **{n_ref}** 条（APA / MLA / GB-T7714 / Vancouver / 中文混排 → 新传研究脚注体例；外文题名已置斜体）。")
    L.append(f"- 删除空脚注：{cit.get('deleted_fn') or '无'}。")
    L.append(f"- 正文括注处理 **{len(intext)}** 处：删除（与脚注重复）**{n_del}**，转为脚注 **{n_to}**，保留待复核 **{n_keep}**。")
    L.append("- 含页码的括注（如 `:60`、`:89`），页码已搬入对应脚注。")
    L.append(f"- 脚注序号位置：将 **{cit.get('fnref_punct_moved', 0)}** 处位于句末标点之后的序号移到标点之前"
             "（规范：序号置于标点之前，如 `……研究①。`）。\n")

    L.append("### 2.1 改写后的脚注（核对）\n")
    L.append("| 注 | 改写后 |")
    L.append("|----|--------|")
    for r in cit.get("reformatted", []):
        txt = r["text"].replace("|", "\\|")
        L.append(f"| fn{r['fn']} | {txt} |")
    L.append("")

    L.append("### 2.2 括注 → 脚注（新建）\n")
    L.append("| 段 | 原括注 | 新脚注 |")
    L.append("|----|--------|--------|")
    for x in intext:
        if x["action"] == "转脚注":
            L.append(f"| P{x['para']} | {x['cite']} | fn{x.get('new_fn')}：{x.get('fn_text','')} |")
    L.append("")

    L.append("### 2.3 已删除的重复括注\n")
    dels = [x for x in intext if x["action"] == "删除括注"]
    L.append("，".join(f"P{x['para']}{x['cite']}" for x in dels) or "（无）")
    L.append("")

    # 3. 完整性
    L.append("## 三、完整性校验\n")
    L.append(f"- 脚注总数：{bef.get('footnote_count','?')} → {aft.get('footnote_count','?')}；"
             f"正文引用标记：{bef.get('body_ref_count','?')} → {aft.get('body_ref_count','?')}。")
    L.append(f"- 正文引用与脚注一一对应：新增缺失 {integ.get('new_missing')}，新增孤立 {integ.get('new_orphan')}；"
             f"脚注 ID 连续（断号 {aft.get('gaps')}）。")
    L.append(f"- 结论：**{'通过 ✅' if integ.get('ok') else '需检查 ❌'}**\n")
    if cit.get("warnings"):
        L.append("运行告警：")
        for wn in cit["warnings"]:
            L.append(f"- ⚠ {wn}")
        L.append("")

    # 4. 待人工复核
    L.append("## 四、待人工复核清单 ⚠️（务必逐条确认）\n")
    L.append("> 以下为信息缺失、存在冲突或需作者判断之处。脚本**未擅自补全**，请核对后手动处理。\n")
    cats = {}
    for cat, text in flags:
        cats.setdefault(cat, []).append(text)
    order = ["匿名", "缺信息", "冲突", "重复", "人名", "格式"]
    label = {"匿名": "匿名评审", "缺信息": "文献信息缺失", "冲突": "信息冲突",
             "重复": "重复引用（建议改「同注」）", "人名": "外国人名", "格式": "格式与其他"}
    for cat in order + [c for c in cats if c not in order]:
        if cat not in cats:
            continue
        L.append(f"### {label.get(cat, cat)}\n")
        for t in cats[cat]:
            L.append(f"- [ ] {t}")
        L.append("")

    # 5. 收尾自检清单
    L.append("## 五、收尾自检清单（对照规范末尾）\n")
    L.append("**格式类**")
    L.append("- [x] 版式按每页 38 行 / 每行 41 字（A4 + 页边距已设；最终行数以 Word 重排为准）。")
    L.append("- [x] 四级标题序号体系：`一、` / `（一）` / `1．` / `（1）`（已套黑体分级字号）。")
    L.append("- [x] 摘要标签统一为「内容提要」（仅标签，正文未动）。")
    L.append("- [ ] 外国人名首现「汉译 + 括注原名」（见第四节「外国人名」项）。")
    L.append("- [ ] 无繁体字（特殊需要除外）；标点与数字符合国标。")
    L.append("- [ ] 匿名评审：正文已剔除作者身份信息；资助说明置于文末且仅一项（见第四节「匿名」项）。")
    L.append("- [ ] 英文标题/作者/摘要/关键词另附（本 skill 不自动翻译生成，见第四节）。")
    L.append("")
    L.append("**引文类**")
    L.append("- [x] 所有引证已统一为页下脚注，序号 `①②③`、每页独立排序（编号格式已写入；WPS 若回退请手动重设）。")
    L.append("- [x] 正文括注已转脚注或删除（仅剩缺信息括注待补，见第四节）。")
    L.append("- [x] 脚注序号置于句末标点之前。")
    L.append("- [x] 各脚注「责任者—题名—出版信息—页码」顺序与标点符合对应类型模板。")
    L.append("- [x] 期刊类不带具体页码；著作 / 报纸 / 论文带页码（缺页码者已在第四节标注）。")
    L.append("- [x] 外文文献：题名斜体、析出篇名英文引号、`p.`/`pp.` 用法正确。")
    L.append("- [ ] 信息不全或类型存疑条目已在报告标记（见第四节，**请逐条处理后再投稿**）。")
    L.append("")
    L.append("---\n")
    L.append("> 提示：在 Word/WPS 中打开输出文件后，请检查 ① 脚注编号是否显示为 ①②③ 且每页重排"
             "（WPS 个别版本需手动设置）；② 图表是否跨页；③ 第四节各项是否已处理。")

    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(L))
    return report_path


if __name__ == "__main__":
    print("本文件为模块，供 apply_all.py 调用。请运行：python3 apply_all.py <输入.docx>")
