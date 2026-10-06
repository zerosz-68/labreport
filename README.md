# labreport

[![CI](https://github.com/zerosz-68/labreport/actions/workflows/ci.yml/badge.svg)](https://github.com/zerosz-68/labreport/actions/workflows/ci.yml)
[![Python](https://img.shields.io/badge/python-3.9%2B-blue)](https://github.com/zerosz-68/labreport)
[![License: MIT](https://img.shields.io/badge/license-MIT-green)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey)](#跨平台与验证)

> **把实验数据填进 Word 实验报告模板的 CLI。**
> 合并单元格不会错位、公式是 **Word 原生公式**、填完自动自检、能转 PDF 供肉眼复核。

```console
$ labreport inspect 报告模板.docx          # 看清表格物理网格（含合并单元格）
$ labreport data check data.json           # 校验手写数据：单位/平均值/离群/公式重算
$ labreport fill 报告模板.docx --map map.json --apply
$ labreport audit 成品.docx --source 报告模板.docx --data data.json
结论: 可以交付（error 0 / warn 0）
```

---

## 目录

- [它解决什么问题](#它解决什么问题)
- [60 秒上手](#60-秒上手)
- [工作流](#工作流)
- [命令一览](#命令一览)
- [三个核心概念](#三个核心概念)
- [文档](#文档)
- [项目结构](#项目结构)
- [跨平台与验证](#跨平台与验证)
- [贡献与许可](#贡献与许可)

---

## 它解决什么问题

| 手工填实验报告的痛点 | labreport 的做法 |
|---|---|
| 表格有**合并单元格**，用 python-docx 按 `cells[r][c]` 填会整体错位 | 从 OOXML 的 `gridSpan`/`vMerge` 还原**物理网格**，按"屏幕上看到的行列"定位 |
| 公式要一个个用 Word 公式编辑器敲 | 内置 **LaTeX → OMML** 转换（不依赖 pandoc），直接产出 Word 原生公式 |
| 抄错一位数字，要人工从头核一遍 | `data check` 自动查：单位归一化、平均值自洽、3σ 离群、公式重算与预期值比对 |
| 填完不知道有没有漏填 / 结构有没有被弄坏 | `audit` 对比模板结构、统计公式、扫占位残留、把数据**逐项回读** |
| 让 AI 直接改 docx，常常把排版改烂 | 默认**只出填写计划**（旧值→新值），确认后才落地；改完还能 `pdf --pages` 出图复核 |

## 60 秒上手

```bash
pip install git+https://github.com/zerosz-68/labreport
labreport doctor          # 环境体检：依赖 / 可选程序 / 中文字体
```

不想装东西也能用（零安装入口）：

```bash
git clone https://github.com/zerosz-68/labreport && cd labreport
python run.py doctor
```

**看一个完整跑通的例子**（自动生成模板与数据，不需要任何真实报告）：

```bash
python examples/demo.py
```

它会造一份带合并单元格的模板，依次跑 `inspect → data check → fill → audit`，并打印每一步的关键输出。

## 工作流

```
手写记录照片 ──agent 视觉──▶ data.json ──data check──▶ 校验通过（0 error）
                                                          │
报告模板.docx ──inspect──▶ 物理网格坐标 ──map.json──▶ fill ──▶ 成品.docx（原生公式）
                                                          │
公式清单.md ──omml──▶ OMML 片段 ────────────────────────────┘
                                                          ▼
                                          audit（自检）──▶ pdf --pages（看图复核）
```

- **谁能做什么**：看照片识别数字是 agent 的活；`labreport` 不看图，只负责确定性硬活与校验。
- **人工确认**：`fill` 默认只打印计划（旧值 → 新值），加 `--apply` 才写文件；全自动场景用 `--auto`。

## 命令一览

| 命令 | 作用 |
|---|---|
| `doctor` | 环境体检（核心依赖 / 可选程序 / 中文字体） |
| `inspect <模板.docx>` | 读结构：段落、**表格物理网格**、合并单元格、嵌套表、公式、图片、占位 |
| `data template` / `data check <data.json>` | 生成数据骨架 / 校验手写数据 |
| `omml <公式.md> [--compare 参照目录]` | LaTeX → Word 原生公式（OMML） |
| `fill <模板> --map <map.json> [--data <data.json>] [--apply\|--auto]` | 按物理网格填值 + 插公式 |
| `audit <成品> [--source <模板>] [--data <data.json>]` | 成品自检（交付前最后一道关） |
| `pdf <成品.docx> [--pages]` | 转 PDF；`--pages` 渲染成 PNG 供视觉复核 |
| `install-skill --for <agent>` | 把适配页装进 Claude / DSH / Codex / Cursor |
| `card list \| show \| new` | 实验知识卡：换实验只换一张卡 |

退出码：`0` 就绪/无 error · `1` 有错 · `2` 用法错误 · `3` 缺可选程序（降级但可用）。

跑 `labreport <命令> --help` 看每个参数。

## 三个核心概念

1. **物理网格**：`{"table":2,"row":1,"col":2}` 指的是**屏幕上看到的**第 2 张表第 1 行第 2 列；落进哪个真实单元格由合并关系决定。这是它能正确填合并单元格的原因。
2. **数据契约**：JSON 里 `value + unit` 按仪表/照片原样填（`1116.12 g`、`774 ms`），CLI 自动归一化到 SI 供公式使用；每个值都带 `source`，校验失败时知道该回看哪张照片的哪一格。
3. **知识卡**：通用流程在 CLI 里，**每个实验特有的公式/表结构/量级易错点**在 `cards/<实验>.md`。换实验只加一张卡，不动代码。

## 文档

| 文档 | 内容 |
|---|---|
| [快速上手](docs/quickstart.md) | 安装、第一条命令、目录约定 |
| [标准工作流](docs/workflow.md) | 从照片到成品报告的 7 步，每步产出与确认点 |
| [命令手册](docs/commands.md) | 9 条命令逐条说明 + 参数 + 退出码 |
| [数据格式](docs/data-format.md) | `data.json` / `map.json` 全部字段与规则 |
| [知识卡](docs/cards.md) | 怎么为新实验写卡、内置卡有哪些 |
| [常见问题](docs/faq.md) | 编码、合并单元格、公式、Word/LibreOffice、隐私、CI |
| [回归测试](tests/README.md) | 用例覆盖与如何本地跑全量 |

## 项目结构

```
labreport/
├─ src/labreport/
│  ├─ cli.py           命令行入口（9 个子命令）
│  ├─ doctor.py        环境体检
│  ├─ docx/            inspect（物理网格）· fill（填写）
│  ├─ omml/            LaTeX → OMML 内置转换
│  ├─ data/            数据校验（SI 归一化 / 统计 / 重算）
│  ├─ audit/           成品自检（结构 / 公式 / 占位 / 回读）
│  ├─ topdf.py         docx → PDF → PNG（目检）
│  ├─ skills/          四份 agent 适配页（共用一份正文）
│  └─ cards/           实验知识卡（内置 + 自建）
├─ examples/demo.py    端到端演示（自动造模板与数据）
├─ tests/              回归测试（fixtures + 一键脚本）
└─ docs/               文档
```

## 跨平台与验证

| 环境 | 覆盖 | 结果 |
|---|---|---|
| GitHub Actions · ubuntu / macos / windows × Python 3.9 / 3.12 | 安装 + 15 个不需要真实 docx 的用例 + 命令冒烟 | **6/6 全绿** |
| Windows 11 + Python 3.12 + Word + 真实实验文件 | 全量 24 个用例（含 fill / audit / pdf→PNG 目检） | **24/24 通过** |

细节与仍未被 CI 覆盖的点（macOS/Linux 上 `pdf` 的真实转换、中文字体）见 [常见问题](docs/faq.md#跨平台)。

## 贡献与许可

- 提 Issue / PR 前请先看 [CONTRIBUTING.md](CONTRIBUTING.md)；改动请跑 `python tests/run_all.py`。
- MIT License，见 [LICENSE](LICENSE)。
