# 新闻传播学中文学术期刊投稿格式 Claude Code Skills

你的论文写完了。田野回来了，理论框架搭好了，结论也站得住脚了。

然后你打开了投稿系统，看到了格式要求。

> *"页下脚注，①②③，每页独立排序。"*
> *"责任者：《题名》，出版地：出版者，年份，第X页。"*
> *"外国人名第一次出现须译成汉语，括注外文原名。"*

你花了三天改格式。编辑部退稿，说脚注体例不符。你又花了两天。

**这个仓库就是为了让这件事不再发生。**

一组 [Claude Code](https://claude.ai/code) Skills，将论文 `.docx` 自动调整为指定期刊的投稿格式。括注变脚注、APA 变本刊体例、序号归位到句号前——改完还给你一份修改报告，列出所有它没把握的地方，绝不替你瞎猜。每个子目录对应一本期刊，相互独立、按需安装。

## 包含的期刊

| 目录 | 期刊 | 封面 | 主要功能 |
|------|------|------|----------|
| [`xwycbyj-formatter/`](./xwycbyj-formatter/) | 《新闻与传播研究》 | <img src="assets/xwycbyj-cover.jpg" width="120"> | 脚注体例统一、括注/APA/MLA/GB → 本刊格式、版面字体标题、修改报告 |

各期刊的安装方法与触发词详见各子目录的 `README.md`。

## 工作原理

采用**代码 + Claude 判断**的混合策略：

1. **结构抽取**（脚本）：读取 `.docx`，打印全部脚注、括注、标题层级、前置信息
2. **引文映射**（Claude 判断）：逐条把各式引文映射到目标格式，生成判断表 `citation_map.py`
3. **写回 + 报告**（脚本）：机械格式（字体/行距/页边距/脚注编号）+ 引文写回，产出新 `.docx` 和 Markdown 修改报告

> **铁律**：转换时只允许重排已有文献信息的顺序与标点，**严禁新增、补全或推断任何文献内容**（作者、年份、页码、出版社等）。信息缺失一律在修改报告中标记待人工复核，绝不擅自猜测。

## 前置依赖

- [Claude Code](https://claude.ai/code)（CLI 版）
- Python 3.9+
- `pip install python-docx lxml`

## 安装

**第一步：克隆仓库**

```bash
git clone https://github.com/wgyang28/xinchuan-journal-formatter.git
cd xinchuan-journal-formatter
```

**第二步：安装所需期刊的 Skill**

按需选择，每本期刊独立安装：

```bash
# 《新闻与传播研究》
ln -s "$(pwd)/xwycbyj-formatter" ~/.claude/skills/xwycbyj-formatter
```

> 符号链接方式推荐：仓库执行 `git pull` 后，本地 skill 自动同步最新版本，无需重新复制。

## 新增期刊

1. 在仓库根目录新建子目录 `<期刊缩写>-formatter/`
2. 参照 `xwycbyj-formatter/` 的结构，替换 `references/` 里的规范文件和 `format_config.json`
3. 在子目录内新建 `README.md`，写明安装方法与触发词
4. 更新本 README 的期刊列表
