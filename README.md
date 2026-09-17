# Agent Pitfalls (Windows)

**AI coding agent 在 Windows 上干活时反复踩到的坑，以及踩完之后定下的规矩。**

7 个 Agent Skill（约定格式：每个目录一份 `SKILL.md`，frontmatter 的 `description` 决定 agent 何时加载它）。
每条都来自真实事故复盘——不是"最佳实践"式的通用建议，而是
"这一天白干了，因为 X" 之后写下来的对策。中文为主，命令示例面向 Windows + Git Bash。

## 装机

把 `skills/` 下你想用的目录整个复制进你 agent 的技能目录即可（Claude Code `~/.claude/skills`、
Codex `~/.codex/skills`、或任何读 `~/.agents/skills` 的客户端）：

```bash
cp -r skills/* ~/.claude/skills/
```

无需重启；agent 会在匹配的场景自动读取。也可以直接当普通 markdown 文档看。

## 这 7 个的关系

不是 7 个并列的清单，是两层：

- `agent-exec-discipline` 是父层——它给出根因三分法（**忽视了什么 / 遗漏了什么 / 哪个行动出错**）。
- 其余 5 个坑清单按这个三分法逐条实例化；`docx-mathtype-extract` 是纯技术叶子。

只想读一条的话，读 `agent-exec-discipline`。

## 清单

| Skill | 一句话 |
|---|---|
| [agent-exec-discipline](skills/agent-exec-discipline/SKILL.md) | 元规则：根因三分法、开工五问、验证优先于断言、简化优先于手工构造、重复错误止损 |
| [sandbox-git-gh-pitfalls](skills/sandbox-git-gh-pitfalls/SKILL.md) | 沙箱里 git/gh 的坑：unborn 分支与 ref 消失、GCM 取 token、push OOM 与 postBuffer、`gh pr merge` 静默失败、push exit 128 无输出时改走 Contents API |
| [edit-script-pitfalls](skills/edit-script-pitfalls/SKILL.md) | 编辑精确性：heredoc/`python -c` 吃反斜杠与反引号、并行 Edit 静默丢改动、日志追加变覆盖、PowerShell BOM 乱码、Git Bash 盘符路径 |
| [file-parse-tooling-pitfalls](skills/file-parse-tooling-pitfalls/SKILL.md) | 解析与工具：docx 公式在 `<m:t>` 不在 `<w:t>`、PDF 别手解、别假设工具存在、Glob 空结果复核、无外网时不要重试 |
| [numeric-solver-discipline](skills/numeric-solver-discipline/SKILL.md) | 数值求解：自适应步进点不能当约束检查点、扫参找最优的上下限纪律、"可复现"该怎么证明 |
| [docx-final-surgery](skills/docx-final-surgery/SKILL.md) | 对已定稿 Word 做定点手术：OMML 公式批量编号、多子图图题清理、身份信息残留清除、逐字节回读 + 真渲染验证 |
| [docx-mathtype-extract](skills/docx-mathtype-extract/SKILL.md) | MathType/OLE 公式取数：docx 文本层 → WMF 字节扫描 → `document.xml` 交叉校验的三级回退，附可运行脚本 |

## 为什么值得读

共同点是**失效模式看不见**：命令退出码 0 但 PR 没合、Edit 返回 "Successfully edited" 但改动少了、
文件"改好了"其实是整段覆盖、公式在 Word 里显示正常但 `python-docx` 读不到。
这些不会出现在任何官方文档里，只能靠别人替你交过学费。

## 约定

- 每个 `SKILL.md` 的 frontmatter 只有 `name` 和 `description`；`description` 里塞满触发词，
  因为 agent 就是靠它决定要不要读正文。
- 正文结构（实测逐文件核对过）：`sandbox-git-gh-pitfalls`、`edit-script-pitfalls`、
  `file-parse-tooling-pitfalls` 的**全部**条目都是 **现象 → 根因 → 对策**；
  `numeric-solver-discipline` 只有第 1 条（最贵的那课）如此，其余是规则清单；
  `agent-exec-discipline` 与两个 docx 文件是流程/规则文体。对策一律给可直接复制的命令。
- 示例里的路径、端口、文件名都是占位或中性化写法；执行前请按自己机器替换。
- 每处"已付过的学费"都对应一次真实故障，日期保留在正文里；不接受没有事故来源的条目。

## 贡献

PR 欢迎，请保持三件事：一条坑对应一次可复述的真实故障；给出判定命令（怎么确认就是它）；
写清"别做什么"，因为错误路径通常比正确路径更常触发。

LICENSE: MIT
