# labreport

[![CI](https://github.com/zerosz-68/labreport/actions/workflows/ci.yml/badge.svg)](https://github.com/zerosz-68/labreport/actions/workflows/ci.yml)
[![Python](https://img.shields.io/badge/python-3.9%2B-blue)](https://github.com/zerosz-68/labreport)
[![License: MIT](https://img.shields.io/badge/license-MIT-green)](LICENSE)

通用实验报告 CLI。目标：**任何 agent、任何系统、任何实验**都能用同一套流程填报告。

## 为什么是 CLI 而不是"胖 skill"

| 代 | 形态 | 弱点 |
|---|---|---|
| 1 代 | 胖 skill，依赖别人的 skill | 换环境缺零件 |
| 2 代 | 胖 skill，自带脚本 | 依赖靠人装、路径靠工具解析、流程靠模型自觉 |
| **3 代** | **CLI 内核 + 薄 skill + 知识卡（本项目）** | 前期要多写代码 |

CLI 负责确定性硬活；agent 负责需要"理解"的活（看手写照片、决定数据放哪）；
知识卡沉淀每个实验的专业内容。

## 安装

```bash
pip install git+https://github.com/<you>/labreport     # 或本地开发：pip install -e .
labreport doctor                                       # 环境体检
```

**零安装兜底**（目标环境不方便装包时）：

```bash
python run.py doctor
python run.py inspect 模板.docx
```

## 命令

| 命令 | 状态 | 作用 |
|---|---|---|
| `labreport doctor` | 完成 | 环境体检：Python 包、外部程序、中文字体（退出码 0 就绪 / 3 缺可选 / 1 缺核心） |
| `labreport inspect <模板.docx>` | 完成 | 读结构：段落、**表格物理网格**、合并单元格、嵌套表、公式、图片、占位 |
| `labreport omml <公式.md> [--compare 对照片段目录]` | 完成 | **LaTeX → Word 原生公式（OMML）**，内置转换，不依赖 pandoc |
| `labreport data check <data.json>` | 完成 | 校验手写数据：**按 unit 自动归一化 SI**、平均值自洽、3σ 离群、公式重算与预期值比对、存疑标注 |
| `labreport data template` | 完成 | 输出 data.json 骨架（给 agent 照着填） |
| `labreport fill <模板> --map <map.json> [--data <data.json>] [-o 成品] [--apply\|--auto]` | 完成 | 按**物理网格坐标**填文本/追加/清空 + 插入原生公式；**默认只出计划**，`--apply`/`--auto` 才落地 |
| `labreport audit <成品> [--source <模板>] [--data <data.json>]` | 完成 | 自检：结构对比、公式健康（空公式 / LaTeX 残留）、占位残留、**数据逐项回读**；exit 1 = 不可交付 |
| `labreport pdf <成品.docx> [--pages]` | 完成 | 转 PDF（Word COM 或 LibreOffice）目检；`--pages` 再把每页渲染成 PNG 供视觉检查；无转换器时优雅降级（退出码 3） |
| `labreport install-skill --for <agent>` | 完成 | 把适配页装进 Claude / DSH / Codex / Cursor（`--print` / `--dry-run` / `--list` / `--force`） |
| `labreport card list \| show <实验> \| new <实验>` | 完成 | **实验知识卡**：内置卡 + 自建卡（`~/.labreport/cards`，同名覆盖；`LABREPORT_CARD_DIR` 可改） |

## 工作流（三个文件串起来）

```
照片 ──agent 视觉──> data.json ──data check──> 校验通过
                                                 │
模板.docx ──inspect──> 物理网格坐标 ──map.json──> fill ──> 成品.docx（原生公式）
                                                 │
公式清单.md ──omml──> OMML 片段 ──────────────────┘
```

## 已验证（回归样例：扭摆法测转动惯量，全部真实文件）

| 验证项 | 结果 |
|---|---|
| `omml` 28 条真实公式 vs pandoc 参照片段 | **28/28 结构逐字一致** |
| `omml` 语料外公式（求和/积分/平均值/n 次根/不确定度传递） | 12/12 正常解析，无退化 |
| 生成片段插入真实模板并重载 | 通过（分数/根号/求和结构完整） |
| `inspect` 合并单元格还原 | 通过（模板 9 合并；成品 4 张嵌套表与 32 公式正确归属） |
| `data check` 从原始数据重算 10 个目标量 | 与报告值相差 **0.003%–0.26%**（差异来自报告用了四舍五入的中间值） |
| `data check` 负样本（D 少一位 + 周期抄错） | 8 error + 1 warn，明确指出「平均值不自洽：写了 1556，按 6 次算是 1573.33」，退出码 1 |
| `fill` 按物理网格填 7 处 + 插 2 条公式 | **7/7 文本命中**（含合并单元格 r2c5/r2c7 等），OMML 2 条，表结构 6×8/9 合并保持不变 |
| `fill` 默认不落地 | 通过（只出计划、打印旧值→新值，不写文件） |
| `audit` 自检真实【已填写】报告 | **26/26 数据命中、0 error → 判定"可以交付"**；表几何 7×1 / 6×8 未变，公式 0→32，无 LaTeX 残留 |
| `audit` 自检空白模板 / 半填产物 | 2/26、6/26 命中 → 正确报出"大量数据未填" |
| `pdf` 目检真实报告（Word COM → PDF → PNG） | 通过：封面、表格页、**公式页**渲染正常（32 条原生公式全部由 Word 正确排版），中文与表格边框无破版；默认受限环境下也能跑（首次启动 Word 需放开文件权限——它要创建自己的用户配置，之后就绪） |
| `install-skill` 四份适配页 | 通过：Claude/Cursor 写入幂等（重复安装提示"已是最新"）；Codex 的 `AGENTS.md` 用标记块替换，连装两次仍只有 1 个块 |
| 适配页生效验证 | 把 Codex 版装到 `_test/AGENTS.md` 后，harness **立即把它当作该目录的指令加载**（已清理该测试产物） |

> 打包（`pip install -e .`）在开发用的内置 Python 上无法验证——该发行版没有
> setuptools 与 pip 引导；在正常 Python 环境按上面的安装命令即可。

## 退出码约定

- `0` 就绪 / 无 error（可交付）
- `1` 缺核心依赖、文件或参数错误、校验出 error、定位失败
- `2` 命令行用法错误（argparse 保留，业务不使用）
- `3` 核心齐全但缺可选程序（功能降级，仍可用）

## 知识卡（换实验只换一张卡）

```bash
labreport card list                      # 看有哪些卡（内置 + 自建）
labreport card show torsion-pendulum     # 读卡：公式 / 表结构 / 量级易错点 / 参考数据
labreport card new 用单摆测重力加速度     # 没卡就用模板建一张，做完沉淀回去
```
通用流程在 CLI 里，实验知识在卡里；卡目录可用 `LABREPORT_CARD_DIR` 指到项目内，随项目走。

## 回归测试

```bash
python tests/run_all.py                     # 不依赖 docx，跑一半以上用例
python tests/run_all.py --docx-dir <目录> --frags-dir <目录>   # 全量
```
最近一次（Windows 11 + Python 3.12 + Word COM）：**24 通过 / 0 失败 / 0 跳过**，详见 [tests/README.md](tests/README.md)。

## 跨平台说明

**已实测（CI 证据在 Actions 页，每次推送自动跑）**

| 环境 | 覆盖 | 结果 |
|---|---|---|
| GitHub Actions · ubuntu / macos / windows × Python 3.9 / 3.12 | 安装 + 15 个不需要真实 docx 的用例 + 命令冒烟 | **6/6 全绿** |
| Windows 11 + Python 3.12 + Word + 真实实验文件 | 全量 24 个用例（含 fill / audit / pdf→PNG 目检） | **24/24 通过** |

- 代码是跨平台写的：路径用 `os.path`/`expanduser`、程序探测用 `shutil.which`、平台差异用 `sys.platform` 分支；
- **非 UTF-8 控制台**（英文版 Windows 的 cp1252）也不会崩：`cli.py` 启动时把管道输出切成 UTF-8，终端输出保留原编码但把不可编码字符降级为 `?`（这是从 CI 的 Windows 作业里抓出来的真实 bug，已修）；
- 两点平台差异仍需注意：
  1. `pdf`：Windows 走 Word COM，macOS / Linux 走 `soffice --headless`（CI 只验证"没有转换器时优雅降级为退出码 3"，**真实转换未在 macOS/Linux 上验证**）；
  2. 中文字体：macOS / Linux 建议装 Noto Sans CJK / Source Han Sans，否则 PDF 与图表中文可能缺字（`doctor` 会提示）。
- 打包：CI 里 `pip install -e .` 在三个系统都能装（含 Python 3.9）；开发用的内置 Python 因缺 setuptools 无法本地验证，故提供 `run.py` 零安装入口。
