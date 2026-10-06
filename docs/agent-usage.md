# 在 agent 里使用

`labreport` 的设计前提就是**给 agent 用**：

- **agent 负责**：看照片识别手写数字、判断哪个值放哪个格、跟你确认、看图检查排版；
- **CLI 负责**：解析 OOXML（含合并单元格）、生成原生公式、按坐标写入、数值校验、结构自检、转 PDF/PNG。

两者之间靠一页**适配说明**（薄 skill）连接：装上它，agent 就知道**什么时候**该调用 `labreport`、**怎么调用**。

---

## 1. 一行安装

```bash
labreport install-skill --for claude      # 或 dsh / codex / cursor
```

还没装 `labreport` 时（克隆了仓库）：

```bash
python run.py install-skill --for dsh
```

| `--for` | agent | 落点 | 形态 |
|---|---|---|---|
| `claude` | Claude Code / Claude 桌面 | `~/.claude/skills/labreport/SKILL.md` | YAML frontmatter（description 决定何时触发） |
| `dsh` | DeepSeek Harness | `$DSH_HOME/skills/labreport/SKILL.md` | 同上 |
| `codex` | Codex CLI | `~/.codex/AGENTS.md` | **标记块**（`<!-- labreport:begin/end -->`），不碰你已有的内容 |
| `cursor` | Cursor | `.cursor/rules/labreport.mdc` | MDC frontmatter（`globs: ["**/*.docx", …]`） |

查看落点、预览内容、装到别处：

```bash
labreport install-skill --list                  # 四种 agent 的默认落点
labreport install-skill --for dsh --dry-run     # 只说要写哪里、不落盘
labreport install-skill --for cursor --print    # 打印全文，自己粘到任何 agent
labreport install-skill --for claude --dest ./my/skills/labreport/SKILL.md
```

**幂等**：内容一样时再装一次会提示"已是最新，无需改动"；`--force` 才强制覆盖。
四份适配页共用同一份正文（`src/labreport/skills/_body.md`），改一处、四处同步。

> DSH 会扫描 `$DSH_HOME/skills/<名称>/SKILL.md`（本机是 `C:\Users\noom\.dsh\skills\`）；
> Codex 读 `AGENTS.md` 作为目录级指令——这一点实测过：装到某个目录的 `AGENTS.md` 后，
> agent 立刻按它执行。

---

## 2. 装完怎么用

直接用自然语言，不用记命令：

> 用 labreport 把实验报告填好：模板是 `报告模板.docx`，手写数据在 `照片/` 里。

更具体的例子（交给 agent 的提示词可以这么写）：

> 按 `扭摆实验报告模板.docx` 填报告。数据是 `img1.jpg`~`img10.jpg` 里的手写记录。
> 填之前先 `labreport card show torsion-pendulum` 看有没有现成的实验知识卡；
> 数据要落成 `data.json` 并用 `labreport data check` 校验到 0 error；
> 出计划给我确认后再落地；最后 `labreport audit` 和 `labreport pdf --pages`。

## 3. agent 会走的流程

适配页里写死了顺序（"铁律"），所以不同 agent 的行为是一致的：

```
0  labreport doctor                     核心依赖缺失就停下，不绕过
   labreport card show <实验>           有知识卡先读卡（公式/表结构/量级易错点/参考数据）
1  labreport inspect 模板.docx          拿到物理网格：几行几列、哪些是合并格
2  读照片 → data.json（每个值带 source）
   labreport data check data.json       有 error 就回看照片，禁止猜
3  写 map.json（值/公式落到 表N 行R 列C）
4  labreport fill ...                   ★先只出计划 → 给你看 → 你确认 → --apply
5  labreport audit 成品 --source 模板 --data data.json   必须 0 error
6  labreport pdf 成品 --pages           agent 自己看图确认排版
7  labreport card new <实验>            没卡就把这次的经验沉淀成卡
```

**人在环上（两个确认点）**：

| 确认点 | 默认行为 | 想跳过时 |
|---|---|---|
| 填写前 | `fill` 只打印「旧值 → 新值」计划，**不写文件** | 对 agent 说"全自动"，它会用 `--auto` |
| 交付前 | `audit` 必须 0 error 才允许声称完成 | ——（这是硬门槛，不该跳过） |

## 4. 适配页里到底写了什么

`install-skill --for <agent> --print` 能看到全文，共六节：

| 节 | 内容 |
|---|---|
| 0 铁律 | 五条顺序约束：先 `doctor`、看照片是 agent 的活、数据必须先校验、默认不写文件、交付前必须自检 |
| 1 命令表 | 九条命令 + 退出码语义 |
| 2 标准流程 | 上面那条 0→7 的流水线 |
| 3 `data.json` 示例 | 最小可用样例 + 三个要点（不必自己换算单位、`source` 必填、`targets` 顺序求值） |
| 4 `map.json` 示例 | 物理网格坐标、四类动作、可用占位符 |
| 5 排错表 | 七种常见现象的处置（平均值不自洽、超出网格、LaTeX 残留、`pdf` 退出码 3…） |
| 6 边界 | 只按模板填、不编造数据、未经确认不写文件 |

它**刻意不写实现细节**（怎么解析 OOXML、怎么生成 OMML）——那些在 CLI 里，agent 不需要知道，也不该自己动手重写。

## 5. 验证是否装好

```bash
labreport install-skill --for dsh            # 再装一次
# 期望输出：已是最新，无需改动
```

- **Claude / DSH**：确认 `<落点>/SKILL.md` 存在，且开头有 `name:` / `description:` 两行；
- **Codex**：确认 `~/.codex/AGENTS.md` 里**只有一个** `labreport:begin`…`labreport:end` 块（重复安装不会堆积）；
- **Cursor**：确认 `.cursor/rules/labreport.mdc` 存在且首行是 `---`。

装完**新开一个会话**（多数 agent 只在会话开始时加载指令）。

## 6. 常见问题

**Q：可以让 agent 直接改 docx，不用这个 CLI 吗？**
能，但结果通常更差：合并单元格容易错位、公式容易变成图片或纯文本、填完没有客观校验。本项目就是为这些坑准备的。

**Q：一个机器上装了多个 agent，会冲突吗？**
不会。四份适配页各自落点独立；正文同一份，装哪个都行，也可以四个都装。

**Q：agent 说"环境缺依赖"怎么办？**
按它说的 `pip install` 装核心依赖（`python-docx`/`lxml`/`Pillow`）。缺 LibreOffice 只是不能出 PDF 目检，不影响填写。

**Q：agent 不按流程走 / 跳过了校验？**
把这一句直接发给它：**"先跑 `labreport data check` 到 0 error，再 `fill`，最后 `audit` 必须 0 error"**。也可以在 `AGENTS.md` / 项目规则里加一条硬性要求。

**Q：我用的 agent 不在列表里（比如别的 IDE 插件）？**
`labreport install-skill --for cursor --print` 把全文打出来，粘到那个 agent 的指令文件（通常是 `AGENTS.md`、`CLAUDE.md`、`.rules` 之类）即可——它本质就是一段 Markdown。
