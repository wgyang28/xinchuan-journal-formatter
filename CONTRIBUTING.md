# 贡献指南

欢迎为本仓库贡献新的期刊 Skill，或改进现有功能。

## 贡献方式

### 方式一：新增期刊 Skill

1. **Fork** 本仓库，克隆到本地
2. 在仓库根目录新建子目录，命名规则：`<期刊缩写>-formatter/`
3. 参照 `xwycbyj-formatter/` 的结构搭建：

   ```
   <期刊缩写>-formatter/
   ├── SKILL.md                  # Claude Code skill 定义（name/description/触发词）
   ├── README.md                 # 安装方法、触发词、自动处理内容说明
   ├── references/
   │   ├── <期刊名>_格式与注释规范.md   # 期刊官方规范（需注明来源）
   │   └── format_config.json           # 版面参数
   └── scripts/
       ├── citation_map.py       # 示例引文映射（用虚构文献，不含真实稿件内容）
       └── ...                   # 可直接复用现有脚本（apply_all/format/citations 等）
   ```

4. 在 `citation_map.py` 中**只使用虚构文献**作示例，不提交真实稿件内容
5. 在根目录 `README.md` 的期刊列表中添加一行（含封面图）
6. 提交 **Pull Request**，PR 描述中注明：
   - 期刊全称与主办单位
   - 格式规范文档来源（官网链接或注明"公开发布"）
   - 已在本地测试过的简要说明

### 方式二：改进现有 Skill

- 修复 bug → 提 Pull Request，描述问题与修复方式
- 改进脚本 → 确保改动不破坏现有期刊的处理流程
- 更新期刊规范 → 在 PR 中说明规范版本与生效日期

### 方式三：提 Issue

- **报告 bug**：使用 Bug Report 模板，提供期刊名、复现步骤、错误信息
- **申请新期刊**：使用 New Journal Request 模板，说明期刊名称与规范来源

## 铁律（所有贡献必须遵守）

> 转换时只允许重排已有文献信息的顺序与标点，**严禁新增、补全或推断任何文献内容**。信息缺失一律标记待人工复核，绝不擅自猜测。

- 不提交真实稿件内容（正文、脚注、田野笔记等）
- `citation_map.py` 示例只用虚构文献
- 不在代码里硬编码任何个人信息

## 开发环境

```bash
pip install python-docx lxml
python3 --version  # 需要 3.9+
```

测试：
```bash
python3 scripts/footnote_ops.py selftest <任意.docx>
python3 scripts/apply_all.py <测试.docx>
```

## 代码风格

- Python 文件顶部保留 `# -*- coding: utf-8 -*-`
- 函数/模块说明用中文，保持与现有代码一致
- 不产生 JSON 中间文件；数据走内存，最终只落地 `.docx` 和 `_修改报告.md`
