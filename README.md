# 新闻传播学中文学术期刊投稿格式 Claude Code Skills

一组 Claude Code Skills，每个子目录对应一本期刊，将论文 `.docx` 自动调整为该刊投稿格式。

## 包含的期刊

| 目录 | 期刊 | 说明 |
|------|------|------|
| [`xwycbyj-formatter/`](./xwycbyj-formatter/) | 《新闻与传播研究》 | 脚注体例统一、括注/APA/MLA/GB → 本刊格式、版面字体标题 |

## 安装方法

将所需期刊的子目录复制或链接到 `~/.claude/skills/`：

```bash
# 方法一：符号链接（推荐，修改仓库即生效）
ln -s /path/to/journal-skills/xwycbyj-formatter ~/.claude/skills/xwycbyj-formatter

# 方法二：直接复制
cp -R xwycbyj-formatter ~/.claude/skills/
```

## 依赖

```bash
pip install python-docx lxml
```

## 使用

在 Claude Code 中输入触发词（见各 skill 的 `SKILL.md` description 字段），或直接说"帮我把这篇论文调成新传研究的格式"。

## 新增期刊

1. 在仓库根目录新建子目录 `<期刊缩写>-formatter/`
2. 复制 `xwycbyj-formatter/` 结构，修改 `references/` 里的规范文件与 `format_config.json`
3. 更新本 README 的期刊列表
