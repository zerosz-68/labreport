# 命令手册

九个命令，按典型使用顺序排列。任何命令都可以 `--help` 看参数。

```bash
labreport --version        # 版本
labreport --help           # 命令列表
```

## 退出码约定（所有命令通用）

| 码 | 含义 |
|---|---|
| `0` | 就绪 / 无 error（可交付） |
| `1` | 缺核心依赖、文件或参数错误、校验出 error、定位失败 |
| `2` | 命令行用法错误（argparse 保留，业务不使用） |
| `3` | 核心齐全但缺可选程序（功能降级，仍可用） |

---

## 1. `doctor` —— 环境体检

```bash
labreport doctor [--json]
```

检查 Python 版本、三个核心依赖、六个可选程序（LibreOffice / Pandoc / draw.io / xelatex / pdftoppm / Tesseract）、中文字体提示。
**任何任务开始前先跑它**：缺核心依赖就停下装，缺可选程序只是少功能。

```console
$ labreport doctor ; echo $?
...
缺   LibreOffice      转 PDF 目检（没有则跳过排版目检）
3
```

---

## 2. `inspect` —— 读 Word 结构

```bash
labreport inspect <模板.docx> [--json] [--max-text N] [--formula-sample N]
```

输出内容：

- 段落数（含非空）、内嵌图片数、OMML 公式数；
- 每张表的 `行 × 列`、**逻辑单元格 → 真实单元格**、合并数、空白数、占位、本表公式、含图、嵌套表；
- 逐行的物理网格预览（合并格标 `⇢N列` / `⇣合并`，空格标 `〔空〕`，公式标 `〔公式〕`，嵌套表标 `〔嵌套表×N〕`）；
- 公式样例文本。

```console
$ labreport inspect 模板.docx --max-text 14
  ── 表2: 6 行 × 8 列 | 逻辑单元格 17 → 真实单元格 17 | 合并 9 | 空白 9 | 占位 0 | 本表公式 0 | 含图 0 | 嵌套表 0
    r1  | 实验名称 | 〔空〕⇢5列 | 成绩 | 〔空〕
    r2  | 班级 | 〔空〕 | 学号⇢2列 | 〔空〕 | 姓名 | 〔空〕⇢2列
```

`--json` 输出完整结构（含每个物理格到锚点的映射），适合程序消费。

---

## 3. `data` —— 数据骨架与校验

```bash
labreport data template [-o data.json]     # 输出骨架模板
labreport data check <data.json> [--json]  # 校验
```

`check` 做的事：

| 检查 | 说明 |
|---|---|
| 单位归一化 | 按 `unit` 换算到 SI（`g→kg`、`mm→m`、`ms→s`…），表达式里直接用 SI 值 |
| 平均值自洽 | 声明的 `mean` 与按 `values` 算出的平均是否一致（容差由 `precision` 决定） |
| 离群 | 超过 3σ 的测量值会被点名 |
| 样本量 | 少于 3 次会给提示（A 类不确定度不可靠） |
| 公式重算 | 逐条求值 `targets`，与 `expected` 比对；超出 `tolerance` → warn，超出 5 倍 → error |
| 可追溯性 | 缺 `source` 的项会提示（校验失败时不知道回看哪张照片） |

`--json` 给出机器可读的完整报告（含每项算出的值），便于流水线判断。

---

## 4. `omml` —— LaTeX 转 Word 原生公式

```bash
labreport omml <公式.md> [-o 输出目录] [--compare 参照目录] [--stdout N]
```

- 输入：每行一条公式的清单（可带 `$...$`，`#` 开头为注释）；
- 输出：`fNN.xml`（每条一个 `<m:oMath>` 片段）+ `frags.json`；
- `--compare`：与一个已有 OMML 片段目录逐条做**结构比对**（我们用它对着 pandoc 的输出做过回归，28/28 逐字一致）；
- `--stdout N`：只打印第 N 条的 XML。

**内置转换器，不需要 pandoc。** 支持：`\frac`、`\sqrt`（含 `\sqrt[n]{}`）、`^`/`_`（可同时）、`\left(...\right)`、普通括号（会生成 `m:d` 定界符）、`\sum`/`\int`（带上限）、`\bar`/`\vec`/`\hat`/`\dot`、希腊字母、常用运算符、`\mathrm`/`\text`（正体）、`\,` 等空白。

```bash
$ labreport omml formulas.md -o out --compare pandoc_frags
转换完成：28 条 -> out
与 pandoc 对照：28/28 条结构完全一致
```

---

## 5. `fill` —— 按物理网格填写

```bash
labreport fill <模板.docx> --map <map.json> [--data <data.json>] [-o 成品.docx]
                [--apply | --auto] [--plan <计划.json>]
```

| 参数 | 说明 |
|---|---|
| `--map` | **必填**，填写映射（见 [数据格式](data-format.md#mapjson)） |
| `--data` | 可选，供 `{{targets.X}}` 等占位符取值；也可写在 map.json 的 `data` 字段 |
| `-o/--out` | 输出文件，默认 `<模板名>-filled.docx` |
| `--apply` | 真正写文件（**不加则只出计划**） |
| `--auto` | 等同于 `--apply`（全自动场景用） |
| `--plan` | 把填写计划另存为 JSON（应用模式下默认会写 `<out>.plan.json`） |

四种动作：`set`（替换，保留段落属性与首个 run 的字体）、`append`（追加段落）、`clear`、`formula`（插入原生公式，可用 `"append": true` 追加）。
定位方式：`{table,row,col}` 物理坐标，或 `{find, occurrence}` 文字锚点。**只出计划时绝不会碰文件。**

---

## 6. `audit` —— 成品自检

```bash
labreport audit <成品.docx> [--source <模板.docx>] [--data <data.json>] [--tolerance R] [--json]
```

| 检查 | 级别 |
|---|---|
| 与 `--source` 对比：表格数量、每表 `行×列`、分节数 | 尺寸变了 → error；合并数变了 → warn |
| 公式：数量、空公式、**LaTeX 残留**（正文里还有 `\frac`/`^{` 说明没转成原生） | error |
| 占位残留：`{{...}}`、`—`、`____`、`（ ）`、`【】` | warn |
| 数据回读：`data.json` 里每个量/均值/目标值能否在成品里找到（识别 `1.4091×10⁻³` 这种写法；容差按各项精度与 `tolerance`） | 找不到 → warn |

`--tolerance` 是数值命中的相对容差（默认 `0.002`），比它更大的差异才判为"没找到"。

---

## 7. `pdf` —— 转 PDF 与出图复核

```bash
labreport pdf <成品.docx> [-o out.pdf] [--pages] [--outdir 目录] [--dpi 110] [--first 1] [--last N]
```

- 转换器优先级：**Microsoft Word（COM，仅 Windows）→ LibreOffice（`soffice --headless`，跨平台）**；
- 都没有时**优雅降级**：打印安装建议，退出码 `3`，不影响其它命令；
- `--pages`：用 `pdftoppm` 把每页渲染成 PNG，交给视觉能力逐页检查（中文、公式、表格不破版）。

> 首次调用 Word COM 需要放开工作区外写入（Word 要创建自己的用户配置目录）；之后就绪，受限环境也能跑。

---

## 8. `install-skill` —— 装进 agent

```bash
labreport install-skill [--for claude|dsh|codex|cursor] [--dest 路径]
                        [--print | --dry-run] [--force] [--list]
```

| agent | 默认落点 | 形态 |
|---|---|---|
| `claude` | `~/.claude/skills/labreport/SKILL.md` | YAML frontmatter |
| `dsh` | `$DSH_HOME/skills/labreport/SKILL.md` | 同上 |
| `codex` | `~/.codex/AGENTS.md` | 标记块（重复安装只替换块内内容） |
| `cursor` | `.cursor/rules/labreport.mdc` | MDC frontmatter（globs / alwaysApply） |

四份适配页共用同一份正文（`src/labreport/skills/_body.md`），所以改一处、四处同步。**幂等**：内容相同会提示"已是最新"。

---

## 9. `card` —— 实验知识卡

```bash
labreport card list
labreport card show <卡名>
labreport card new <卡名> [-o 路径] [--force]
```

- 内置卡随包分发（如 `torsion-pendulum`），自建卡默认在 `~/.labreport/cards/`；
- 同名时**自建卡覆盖内置卡**；
- `LABREPORT_CARD_DIR` 可把卡目录指到项目里（随项目走），也是回归测试用的开关。

---

## 组合用法示例

```bash
# CI / 批处理：全自动跑完并断言
labreport doctor            || echo "缺依赖"
labreport data check data.json
labreport fill 模板.docx --map map.json --data data.json --auto -o 成品.docx
labreport audit 成品.docx --source 模板.docx --data data.json

# 人工复核：先看计划，再落地，再出图
labreport fill 模板.docx --map map.json --data data.json
labreport fill 模板.docx --map map.json --data data.json --apply -o 成品.docx
labreport pdf 成品.docx --pages --dpi 130
```
