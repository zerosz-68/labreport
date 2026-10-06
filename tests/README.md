# 回归测试

一键跑完所有命令并断言退出码（进程内调用 CLI，不经子进程/管道，受限环境也能跑）：

```bash
python tests/run_all.py \
  --docx-dir  "…/扭摆实验-全部资料" \
  --frags-dir "…/扭摆实验-全部资料/脚本与中间文件/omml/frags"
```

## 覆盖范围

| 组 | 用例 |
|---|---|
| 环境与公式 | `doctor`；`omml` 28 条真实公式；**与 pandoc 参照逐条比对（28/28）**；语料外公式不退化 |
| 数据校验 | `data template`；`data check` 正确数据（0 error）；抄错版（8 error + 1 warn，退出码 1） |
| 结构/填写/自检 | `inspect` 模板与成品；`fill` 默认**不写文件**（已核实）；`fill --auto` 落地；`audit` 真实报告判"可以交付"；`audit` 空白模板报"大量未填"；`pdf`（无转换器时降级 3） |
| 适配页与知识卡 | `install-skill --list/--print`（四种 agent）；`card list/show/new` + 重复新建拒绝 |

## 关于 fixtures

- `fixtures/*.json`、`fixtures/*.md` 随仓库提供，**不跑 docx 用例也能验证一半以上**；
- 真实 `.docx`（空白模板 / 已填写报告）属于用户资料，**不入库**：用 `--docx-dir` 指过去即可；
  不给就自动跳过并打印原因；
- 比对用的 pandoc 参照片段同样用 `--frags-dir` 指定。

## 历史结果

| 日期 | 环境 | 结果 |
|---|---|---|
| 2026-10 | Windows 11 + Python 3.12（内置运行时）+ Word COM | 24 通过 / 0 失败 / 0 跳过 |
