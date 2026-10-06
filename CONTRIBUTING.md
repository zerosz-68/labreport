# 贡献指南

欢迎 Issue 与 PR。这是一个小工具，保持简单比堆功能重要。

## 环境

```bash
git clone https://github.com/zerosz-68/labreport
cd labreport
pip install -e .
python tests/run_all.py        # 不需要真实报告，跑一半以上用例
```

想跑全量（含真实 docx 用例）：

```bash
python tests/run_all.py --docx-dir <含 原始资料/ 与 成品/ 的目录> --frags-dir <pandoc 参照片段目录>
```

## 提交前

1. `python tests/run_all.py` 必须 0 失败；
2. 改了 CLI 行为 → 同步更新 `docs/commands.md`；
3. 改了数据格式 → 同步更新 `docs/data-format.md` 与 `labreport data template` 的输出；
4. 面向用户的改动 → 在 `CHANGELOG.md` 的"未发布"里加一行；
5. 中文注释/文档保持一致风格（本项目文档为中文，代码标识符为英文）。

## 代码约定

| 项 | 约定 |
|---|---|
| Python | 3.9+，只用标准库 + `python-docx` / `lxml` / `Pillow` |
| 注释 | 解释**为什么**，不复述代码；踩过的坑要写清楚 |
| 输出 | 任何命令都要能在"缺可选程序"时优雅降级（退出码 3），不要抛栈 |
| 编码 | 不要假设 stdout 是 UTF-8（Windows 英文系统是 cp1252） |
| 依赖 | 不新增第三方依赖，除非有充分理由并经讨论 |
| 子进程 | 输出重定向到文件而不是管道（某些受限环境管道不可用） |

## 新增一个实验的支持

**不需要改代码**——写一张知识卡即可：

```bash
labreport card new <实验名>          # 生成模板
# 填公式 / 表结构 / 量级易错点 / 参考数据
```

如果某个实验需要 CLI 层面的新能力（新动作、新校验），请先开 Issue 说明场景，
附上**脱敏后的**模板片段与一份最小 `data.json`。

## 隐私

- **不要**把真实实验报告（`.docx`）、扫描照片、成绩单提交进仓库；
  `.gitignore` 已默认拦掉 `*.docx` / `*.pdf`；
- 用例数据请使用合成的或彻底脱敏的数值；
- 提交邮箱建议用 GitHub 的 noreply 地址。

## 许可

贡献即表示同意以 [MIT License](LICENSE) 发布。
