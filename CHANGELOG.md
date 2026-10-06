# 更新日志

本项目遵循 [语义化版本](https://semver.org/lang/zh-CN/)。

## [未发布]

### 新增

- `examples/demo.py`：端到端演示，现场生成带合并单元格的模板与演示数据，跑完
  `inspect → data check → fill → audit`，不需要任何真实报告。
- `docs/`：拆出快速上手、标准工作流、命令手册、数据格式、知识卡、常见问题六册。

### 修复

- **非 UTF-8 控制台崩溃**：英文版 Windows（cp1252）与 GitHub Actions 的 Windows runner 下
  `print` 中文会抛 `UnicodeEncodeError` 并中断整个命令。现在 `cli.py` 启动时统一处理输出：
  管道/文件走 UTF-8，终端保留系统编码但把不可编码字符降级为 `?`。

## [0.1.0] - 2026-10

首个版本：把"填实验报告"从一次性 skill 变成可复用的 CLI 内核 + 薄 skill + 实验知识卡。

### 新增

- **CLI 九个子命令**：`doctor` `inspect` `data` `omml` `fill` `audit` `pdf` `install-skill` `card`
- **物理网格定位**：从 `w:gridSpan`/`w:vMerge` 还原合并单元格，按"屏幕上看到的行列"填写
- **内置 LaTeX → OMML**：不依赖 pandoc；与 pandoc 参照产物 28/28 结构逐字一致
- **数据校验**：单位归一化、平均值自洽、3σ 离群、公式重算与预期值比对
- **成品自检**：结构对比、空公式、LaTeX 残留、占位残留、数据逐项回读
- **PDF 目检**：Word COM（Windows）或 LibreOffice，`--pages` 渲染成 PNG 供视觉复核
- **四份 agent 适配页**：Claude / DSH / Codex / Cursor，共用一份正文，安装幂等
- **实验知识卡**：内置 `torsion-pendulum`，支持自建与覆盖
- **回归测试**：`tests/run_all.py` 24 个用例，16 个不需要真实 docx 也能跑
- **CI**：ubuntu / macos / windows × Python 3.9 / 3.12

### 验证

- 公式：与 pandoc 参照 28/28 一致；语料外 12 条不退化
- 数据：从手写数据重算 10 个目标量，与报告值差 0.003%–0.26%；抄错版能抓出 8 error
- 填写：7/7 命中（含合并单元格）并插入 2 条原生公式，表格结构无损
- 自检：真实已填写报告 26/26 数据命中、0 error；空白模板能报出大量未填
