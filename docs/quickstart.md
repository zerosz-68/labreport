# 快速上手

## 1. 安装

```bash
# 方式 A：直接从 GitHub 安装（推荐）
pip install git+https://github.com/zerosz-68/labreport

# 方式 B：克隆后开发安装
git clone https://github.com/zerosz-68/labreport
cd labreport
pip install -e .
```

装完在任何目录都能用 `labreport`。要求 Python 3.9+，核心依赖只有三个：
`python-docx`（读写 docx）、`lxml`（解析 OOXML）、`Pillow`（图片处理），`pip install` 会一并装好。

**方式 C：零安装**（不方便装包的环境，例如只读容器、临时机器）：

```bash
git clone https://github.com/zerosz-68/labreport && cd labreport
python run.py doctor
```

`run.py` 只是把 `src/` 加进 `sys.path` 再调用 CLI，功能与安装后完全一致。

## 2. 先做体检

```console
$ labreport doctor
labreport doctor —— 环境体检
  系统      : Windows
  Python    : 3.12.14（要求 >=3.9）
  [核心依赖]
    OK   python-docx    读写 .docx（必需）
    OK   lxml           解析 OOXML（必需）
    OK   Pillow         图片尺寸/OCR 预处理（必需）
  [可选程序：缺失只降级，不阻塞]
    缺   LibreOffice      转 PDF 目检（没有则跳过排版目检）
    OK   Pandoc           LaTeX→OMML 增强（内置转换器可替代）
    缺   draw.io CLI      流程图导出（没有则交源码）
    OK   xelatex          LaTeX 论文编译（没有则只出 docx）
  字体提示  : Windows 自带宋体/黑体/微软雅黑，通常无需处理
```

退出码：`0` 全齐 · `3` 缺可选程序（能用，只是少功能） · `1` 缺核心依赖（先 `pip install`）。

## 3. 跑一个完整例子

仓库自带一个**不需要任何真实报告**的端到端演示：它现场生成一份带合并单元格的模板和数据，把整条流水线跑一遍。

```console
$ python examples/demo.py
[1/5] 生成演示模板 -> examples/_out/模板.docx
[2/5] inspect：识别物理网格与合并单元格
      表2: 6 行 × 4 列 | 真实单元格 9 | 合并 4
...
[5/5] audit：自检
      数据回读  : 12/12 命中
      结论: 可以交付（error 0 / warn 0）
```

看完这一遍，你就知道每个命令的输入输出长什么样了。

## 4. 用在你自己的报告上

```bash
# ① 看清模板结构（拿到物理网格坐标）
labreport inspect 我的实验报告模板.docx

# ② 生成数据骨架，照着填（或让 agent 读照片后填）
labreport data template -o data.json

# ③ 校验数据（有 error 就回看照片，别猜）
labreport data check data.json

# ④ 写填写映射 map.json（哪个值放哪个格）
#    见 docs/data-format.md 的示例

# ⑤ 先看计划，确认后落地
labreport fill 我的实验报告模板.docx --map map.json --data data.json
labreport fill 我的实验报告模板.docx --map map.json --data data.json --apply -o 成品.docx

# ⑥ 自检 + 出图复核
labreport audit 成品.docx --source 我的实验报告模板.docx --data data.json
labreport pdf 成品.docx --pages
```

## 5. 目录约定（可选）

`labreport` 不强制任何目录结构，但下面这套在多实验场景下比较省事：

```
某个实验/
├─ 原始资料/            照片、模板（只读）
├─ data.json            手写数据的结构化结果（每个值带 source 指向照片）
├─ map.json             填写映射
├─ formulas.md          公式清单（每行一条 LaTeX）
└─ 成品.docx            产出
```

知识卡放在 `~/.labreport/cards/`（或用 `LABREPORT_CARD_DIR` 指到项目里，随项目走）：

```bash
labreport card new 用单摆测重力加速度     # 建卡
labreport card list                       # 看有哪些卡
```

## 6. 在 agent 里使用（推荐方式）

`labreport` 本来就是给 agent 用的：你只要说清"模板在哪、数据（照片）在哪"，剩下由 agent 按适配页的流程调用 CLI。

```bash
labreport install-skill --for claude     # Claude Code / 桌面
labreport install-skill --for dsh        # DeepSeek Harness
labreport install-skill --for codex      # Codex CLI（写入 AGENTS.md 的标记块）
labreport install-skill --for cursor     # Cursor（项目规则 .mdc）
```

装完**新开一个会话**，然后直接说：

> 用 labreport 把实验报告填好：模板是 `报告模板.docx`，手写数据在 `照片/` 里。

agent 会自己走完 `doctor → card → inspect → 读照片写 data.json → data check → 填写计划（等你确认）→ fill → audit → pdf --pages`。
**默认会在填写前停下来给你看计划**；想全自动就说"全自动"，它用 `--auto`。

细节（各 agent 落点、如何验证装好、常见问题）见 [在 agent 里使用](agent-usage.md)。

## 下一步

- 想理解每一步在干什么 → [标准工作流](workflow.md)
- 报错了 → [常见问题](faq.md)
