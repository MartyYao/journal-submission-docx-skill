# 生成方法：md → 投稿版 docx

## 方案定案（2026-08-11）：python-docx 全脚本为唯一正式路径

**正式投稿版一律走 python-docx 生成器**（原路线 B）。pandoc（原路线 A）仅作快速草稿预览，不用于正式交付。
理由：13 个期刊模板格式差异大（仿宋/黑体/华文新魏、固定行距 18 磅、三线开口表、上标引用等），reference-doc 无法覆盖；pandoc 的中文字体/行距/上标/三线表均需脚本后处理，不如一步到位。

## 正式链路（定案）

```
1. 合并正文 md（剔除状态行/内部标注）→ 临时 md
2. python-docx 生成器：读取 templates/<期刊>.md 的样式配置 → 逐元素渲染
   （标题/摘要/正文/图表/参考文献各自指定字体+字号+行距+对齐+缩进）
3. officecli 收尾：清 docx 文件属性（作者/最后保存者，匿名要求）+
   期刊要求的精细调整
4. 验证脚本：三线表（无竖线）/图表编号连续/状态行残留为 0/
   字数（Word 口径）/字体字号抽查
```

## 生成器结构（python-docx）

```python
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
# 期刊样式配置从 templates/<期刊>.md 提取（字体/字号/行距/对齐/缩进/页码/页边距）
# 逐元素渲染：标题(#/##/###) → 指定字体；表格 | 行 → 三线表；正文 → 缩进段落
```

要点：
- 中文字体必须双设：`run.font.name` + `rPr.rFonts w:eastAsia`（宋体/黑体/仿宋/楷体/华文新魏）
- 三线表：顶线/栏目线/底线手动设边框，无竖线；表注/资料来源按模板
- 上标引用：[1] 转 run.font.superscript = True；*、**、*** 上标
- 图片：md 相对路径 → 复制到输出目录 → add_picture（宽度按模板）
- 页码域：footer 插入 PAGE 域（多数期刊要求页码）
- 文件属性匿名化：officecli 或 python 清 core.xml 的 creator/lastModifiedBy

## 验证（生成后必跑，脚本化）

```bash
python3 -m venv /tmp/docxenv && /tmp/docxenv/bin/pip install python-docx latex2mathml mathml2omml lxml
# （主环境 PYTHONPATH 被全局注入时需 env -u PYTHONPATH）
env -u PYTHONPATH /tmp/docxenv/bin/python scripts/validate_submission.py <投稿版.docx> <正文md目录>
```

7 项自动检查（任一 FAIL 必须修复）：HTML 标签残留（<sup>）、三线表仅三线、tblGrid 列宽生效、星号上标、图片数一致、页码域、文件属性匿名。
人工复核：字体抽查（XML eastAsia）、表注与表格不跨页分离、长变量名单行、Word 口径字数。

## 环境

- python-docx：**隔离 venv**（`python3 -m venv /tmp/docxenv` + `pip install python-docx latex2mathml mathml2omml lxml`；hermes 主 venv 的 lxml 损坏、PYTHONPATH 全局注入需 `env -u PYTHONPATH` 清理）
- officecli：仅 OpenXML（docx）编辑；属性清理用 python-docx core_properties 更直接
- pandoc：仅草稿预览用（`pandoc -v` 本机已装则用）
