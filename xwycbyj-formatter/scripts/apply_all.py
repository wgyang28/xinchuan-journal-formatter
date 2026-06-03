#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
apply_all.py —《新闻与传播研究》格式 skill 的一站式执行器。

在生成 citation_map.py 之后运行本脚本，一次性完成：
  ① 机械格式（apply_format.run）
  ② 引文写回 + 序号位置（apply_citations.run，依据 citation_map）
  ③ 修改报告（build_report.write_report）

数据全程走内存，**只产出两个文件**：
  <原名>_新传研究投稿版.docx
  <原名>_新传研究投稿版_修改报告.md
不产生任何 JSON 中间文件。

用法：python3 apply_all.py <输入.docx> [-o 输出.docx]
"""

import os
import sys
import argparse

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import apply_format      # noqa: E402
import apply_citations   # noqa: E402
import build_report      # noqa: E402
import citation_map      # noqa: E402


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("-o", "--out", default=None)
    args = ap.parse_args()
    if not os.path.exists(args.docx):
        print("文件不存在: %s" % args.docx, file=sys.stderr)
        sys.exit(1)

    # ① 机械格式
    fmt = apply_format.run(args.docx, args.out)
    out_docx = fmt["output"]

    # ② 引文写回
    cit = apply_citations.run(out_docx)

    # ③ 修改报告
    report_path = os.path.splitext(out_docx)[0] + "_修改报告.md"
    build_report.write_report(report_path, fmt, cit, citation_map.FLAGS)

    # 摘要
    print("=" * 56)
    print("输出文件：%s" % out_docx)
    print("修改报告：%s" % report_path)
    print("-" * 56)
    print("【格式】")
    for c in fmt["format_changes"]:
        print("  ✓ " + c)
    print("【引文】")
    apply_citations.print_summary(cit)
    print("-" * 56)
    print("⚠ 待人工复核 %d 条，详见修改报告第四节。" % len(citation_map.FLAGS))
    print("=" * 56)


if __name__ == "__main__":
    main()
