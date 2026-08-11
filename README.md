# Journal Submission DOCX Skill

> 面向 CSSCI 期刊投稿的 Word（docx）生成器：从论文正文 md 一键生成符合期刊格式要求的投稿版。
> 13 个中文顶刊模板库驱动 + python-docx 全脚本生成 + WPS 兼容公式 + 自动化验证。

## 功能特性

- **13 刊模板库**：中国工业经济 / 经济研究 / 管理世界 / 金融研究 / 会计研究 / 经济学季刊 / 世界经济 / 数量经济技术经济研究 / 管理科学学报 / 南开管理评论 / 改革 / 中国软科学 + `_通用CSSCI` 兜底；模板要点来源为**期刊官网投稿须知**（模板内存档官网 URL 可追溯）
- **python-docx 全脚本生成**（无 GUI 依赖）：正文 md 只读（源稿零改动），逐元素渲染（标题层级 / 字体字号 / 行距缩进 / 三线表 / 图表 / 参考文献 GB/T 7714 / 页码 / 文件匿名）
- **公式 = Word 原生 OMML（WPS 兼容格式）**：LaTeX → latex2mathml → mathml2omml → 复刻 WPS 转换器输出格式（`m:r` 内 `w:rPr` 字体 + DejaVu Math TeX Gyre），WPS 打开直接显示为可编辑公式；公式字号自动跟随上下文
- **表格六条原则**：数字与星号同行、宽表降字号保单行（纯纵向方案，全软件渲染一致）、与正文等宽、垂直居中、行高自适应、列宽按内容
- **章节分页**：每个一级标题另起一页（LibreOffice 兼容方案）
- **自动化验证**：`validate_submission.py` 7 项检查（HTML 残留 / 三线表 / 列宽 / 上标 / 图片 / 页码 / 属性匿名），任一 FAIL 必须修复后才能交付

## 兼容性（多 Agent 实测）

本技能不是 Hermes 专属——SKILL.md + Python 脚本平台无关，**Claude Code / Kimi Code / Codex / Pi 全部实测可用**（2026-08-11，各 Agent 均成功加载技能并跑通验证脚本 7/7）：

| Agent | 安装方式 | 实测 |
|-------|---------|------|
| Hermes | 复制到 `~/.hermes/skills/productivity/` | ✓ |
| Claude Code | 复制到 `~/.claude/skills/` | ✓ |
| Kimi Code | 复制到 `~/.kimi-code/skills/` | ✓ |
| Codex | 复制到 `~/.codex/skills/` | ✓ |
| Pi | `pi install git:github.com/MartyYao/journal-submission-docx-skill` | ✓ |

## 安装

```bash
# 1. 复制到 Hermes 技能目录（或 hermes skills install）
cp -r journal-submission-docx-skill ~/.hermes/skills/productivity/

# 2. 依赖（隔离 venv 推荐；主环境 PYTHONPATH 被全局注入时需 env -u PYTHONPATH）
python3 -m venv /tmp/docxenv
/tmp/docxenv/bin/pip install python-docx latex2mathml mathml2omml lxml
# pandoc（可选，仅草稿预览）
```

## 使用

触发词：**「生成投稿版」「出投稿版」「投稿docx」「按《XX》格式生成」「终稿 docx」** + 13 个模板期刊名之一。

```bash
# 生成：<正文md目录> <输出目录>（自动命名 投稿版-<期刊名>.docx）
/tmp/docxenv/bin/python scripts/generate_submission.py <正文md目录> <输出目录>

# 验证（生成后必跑）
/tmp/docxenv/bin/python scripts/validate_submission.py <投稿版.docx> <正文md目录>
```

输出：`投稿版-<期刊名>.docx`（成品）+ 格式记录 + 生成日志，写入**用户工作文件夹**（不默认写 Obsidian vault）。

## 目录结构

```
journal-submission-docx-skill/
├── SKILL.md                  # 技能主文档（触发条件 / 流程 / 排版规范 / 已知边界）
├── scripts/
│   ├── generate_submission.py   # 生成器（配置驱动，13 刊模板）
│   └── validate_submission.py   # 验证脚本（7 项自动检查）
├── templates/                # 13 刊模板库 + README 索引（官网 URL 存档）
└── references/
    └── generation-methods.md    # 生成方法细节（排版实现 / 环境 / 验证）
```

> 注：`assets/`（期刊官方原件 PDF/doc）仅存本地、不随本仓库分发——官网/编辑部可下载，模板文件内保留官网 URL。

## 配套技能

| 技能 | 负责内容 | GitHub |
|------|----------|--------|
| `paper-workflow` | 论文全流程编排（§7.3 输出转换调用本技能） | [链接](https://github.com/MartyYao/paper-workflow-skill) |
| `meng-skills` | 中文论文写作/润色（写前清单、去 AI 味） | [链接](https://github.com/MartyYao/meng-skills) |

## v0.1.0 说明

首个发布版。核心能力经《中国工业经济》投稿版全流程实测（58 页、21 表、8 图、6 块公式 + 44 行内公式、9/9 章节分页、验证 7/7）：

- 公式链路定案：WPS 兼容 OMML（分析 WPS 转换器输出复刻）——解决 WPS 渲染空白问题
- 表格纯纵向方案（LibreOffice/Word/WPS 渲染一致）
- 表注并入表格、图片 keep_with_next（图表与标题/注不拆分）
- 章节分页隔离空段方案（LibreOffice 兼容）
