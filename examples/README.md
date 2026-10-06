# 示例

## demo.py —— 端到端演示

```bash
python examples/demo.py
```

它不依赖任何真实数据，会现场：

1. 用 `python-docx` 造一份**带水平合并单元格**的模板（信息表 + 数据表 + 公式区）；
2. `inspect` 读出物理网格，让你看到合并单元格是怎么被标出来的；
3. 生成一份演示 `data.json`（含单位、多次测量、三个重算目标）与 `map.json`；
4. `data check` 校验：单位归一化、平均值自洽、按公式重算并与预期值比对；
5. `fill` 先出计划（**不写文件**），确认后 `--auto` 落地，公式以 **Word 原生公式**插入；
6. `audit` 自检：结构有没有变、公式数对不对、数据能不能在成品里逐项回读。

产物都留在 `examples/_out/`（已 gitignore），可以直接用 Word 打开看排版。

## 想接着做什么

```bash
# 出图看排版（需要 Word 或 LibreOffice）
labreport pdf examples/_out/成品.docx --pages

# 换成你自己的模板
labreport inspect 你的模板.docx
labreport data template -o data.json      # 照着填
# 然后按 docs/data-format.md 写 map.json，再 fill / audit
```
