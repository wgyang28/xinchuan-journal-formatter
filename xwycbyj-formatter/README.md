# xwycbyj-formatter —《新闻与传播研究》投稿格式 Skill

将论文 `.docx` 自动调整为《新闻与传播研究》投稿格式，产出新文件 + Markdown 修改报告。

## 安装

```bash
# 在仓库根目录执行（符号链接，推荐）
ln -s "$(pwd)/xwycbyj-formatter" ~/.claude/skills/xwycbyj-formatter

# 或直接复制
cp -R xwycbyj-formatter ~/.claude/skills/
```

## 使用

在 Claude Code 中说出以下任意触发词，skill 自动启动：

- `"帮我把这篇论文调成新传研究的格式"`
- `"新闻与传播研究投稿格式"`
- `"调整格式 + 投稿"`

Skill 会询问输入文件路径，依次完成结构抽取 → 引文映射 → 写回，**只产出两个文件**：

```
原文件名_新传研究投稿版.docx
原文件名_新传研究投稿版_修改报告.md
```

原文件不会被覆盖。

## 自动处理的内容

| 类别 | 具体项 |
|------|--------|
| 版面 | A4 + 页边距，每页约 38 行 × 41 字 |
| 正文 | 宋体/Times New Roman 五号，固定行距 15.6pt，段前段后 0 |
| 标题 | 黑体，四级序号体系（一、／（一）／1．／（1）） |
| 脚注 | ①②③ 每页重排；小五字体；段前段后 0，单倍行距 |
| 摘要标签 | 自动将「摘要」改为「内容提要」 |
| 引文 | 括注/APA/MLA/GB → 本刊页下脚注体例；外文题名斜体 |
| 括注 | 与脚注重复者删除；田野/访谈类转新脚注 |
| 序号位置 | 脚注序号移至句末标点之前（`……研究①。`） |

## 文件结构

```
xwycbyj-formatter/
├── SKILL.md                          # Claude Code skill 定义
├── references/
│   ├── 新闻与传播研究_格式与注释规范.md  # 期刊官方规范（唯一权威依据）
│   └── format_config.json            # 版面参数（字体/字号/行距等）
└── scripts/
    ├── extract_structure.py          # 结构抽取（只读）
    ├── citation_map.py               # 引文映射判断表（每稿由 Claude 生成）
    ├── apply_all.py                  # 一键执行入口
    ├── apply_format.py               # 机械格式模块
    ├── apply_citations.py            # 引文写回模块
    ├── footnote_ops.py               # 脚注 XML 操作库
    └── build_report.py               # 修改报告渲染模块
```
