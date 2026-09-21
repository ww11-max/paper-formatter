# paper-formatter — 学术论文格式修订技能

兼容 Claude Code / OpenAI Codex / OpenClaw 等 Agent Skills 格式的论文格式技能：

1. **修订模式格式调整**：在用户上传的 Word 文档上以修订模式（Track Changes）自动调整格式——
   `inspect_docx.py` 体检 → LLM 对照规范做元素分类映射 → `revise_format.py` 执行，
   全部改动记为可逐条接受/拒绝的 Word 修订，不静默覆盖原文。
2. **中英论文格式规范库**：GB/T 7713.2—2022《学术论文编写规则》结构化清洗稿——
   字号字体表（附录B 表B.1，含 pt 换算与 Word 字体映射）、章节/图表/公式编号、
   量和单位、数字用法、数学式、注释、参考文献，以及 Word 修订模式技术参考与降级方案。

## 安装（任选其一）

| 平台 | 安装位置 |
|---|---|
| Claude Code | `~/.claude/skills/paper-formatter/`（项目级：`<project>/.claude/skills/`） |
| OpenAI Codex CLI | `~/.codex/skills/paper-formatter/` |
| OpenClaw | 工作区 skills 目录（如 `~/clawd/skills/paper-formatter/`），格式同 Agent Skills |
| 其他支持 Agent Skills 规范的加载器 | 任何能读取 `SKILL.md`（YAML frontmatter: name/description）的框架 |

安装后将整个 `paper-formatter/` 目录原样放入上述位置即可，无需构建步骤。

## 环境要求

- **修订模式模块**：Windows + 本机安装 Microsoft Word + Python 依赖 `pywin32`
  （`pip install pywin32`）。文档正被打开时无法处理，请先关闭。
  无 Word 环境的降级方案（python-docx 直改 + 改动对照表，无修订痕迹）
  见 `references/revision-techniques.md` 第 5 节。
- **格式规范库模块**：无任何依赖，纯参考文档。

## 工作流（四步）

```bash
# 1. 体检（只读）：导出逐段格式清单
python scripts/inspect_docx.py 论文.docx --out inspect.json

# 2. 分类映射：对照 references/format-spec-cn-en.md 与 scripts/format_rules.json 的
#    detect_hints，把每段判定为元素类型（title_cn / heading_chapter / body / caption_fig
#    / ref_body 等），写 mapping.json（格式见 revise_format.py 文件头注释）

# 3. 预演
python scripts/revise_format.py 论文.docx --map mapping.json --dry-run

# 4. 执行：全部改动以修订形式写入，输出修订总数与逐段报告
python scripts/revise_format.py 论文.docx --map mapping.json
```

## 说明

- 格式规范以 GB/T 7713.2—2022 为基准；标准中附录 B（字号字体）为**资料性**内容，
  目标期刊有专门投稿模板时以期刊模板优先。
- 期刊/学位论文/报告等不同体裁的编排差异，在映射阶段由 LLM 结合规范判断。
- 标准原文为公开出版的国家标准，本仓库为其格式相关条款的结构化整理与工具化实现。
