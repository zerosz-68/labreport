# labreport —— 实验报告 CLI（薄封装 skill）

> 本 skill 不含实现细节：所有硬活都在已安装的 `labreport` 命令里。
> 你（agent）负责**需要理解的部分**：看照片识别数字、判断数据放哪里、与用户确认。

## 0. 铁律（顺序不能变）

1. 先跑 `labreport doctor`。**核心依赖缺失就停下**，把缺什么、装什么告诉用户，不要绕过。
2. **看照片识别数字是你的活**，`labreport` 不看图；但**不要手写 docx、不要自己拼 OMML 公式**——那是 `fill` 与 `omml` 的活。
3. 数据先落成 `data.json` 并跑 `labreport data check`。**有 error 就回看照片复核，禁止猜**。
4. `fill` 默认**只出计划不写文件**：把「旧值 → 新值」清单给用户确认后，再加 `--apply`。用户明确说"全自动"时用 `--auto`。
5. 交付前必须 `labreport audit` 通过（退出码 0）。有 error 不许声称完成。

## 1. 命令表

| 命令 | 用途 | 何时用 |
|---|---|---|
| `labreport doctor` | 环境体检（Python 包 / 外部程序 / 中文字体） | 每次任务开始 |
| `labreport inspect <模板.docx>` | 读结构：段落、**表格物理网格**、合并单元格、嵌套表、公式、图片、占位 | 填之前必做 |
| `labreport data template [-o data.json]` | 输出 data.json 骨架 | 第一次做某个实验 |
| `labreport data check <data.json>` | 校验：单位归一化、平均值自洽、3σ 离群、公式重算与预期值比对 | 写完数据、填之前 |
| `labreport omml <公式.md> [-o 目录]` | LaTeX → Word 原生公式（OMML），内置转换 | 需要单独产出公式片段时 |
| `labreport fill <模板> --map <map.json> [--data data.json] [-o 成品] [--apply\|--auto]` | 按**物理网格坐标**填文本/追加/清空 + 插入原生公式 | 填写报告 |
| `labreport audit <成品> [--source <模板>] [--data data.json]` | 自检：结构破坏 / 空公式 / LaTeX 残留 / 占位 / 数据回读 | 交付前必做 |
| `labreport pdf <成品.docx> [--pages]` | 转 PDF（Word 或 LibreOffice）；`--pages` 再渲染成 PNG 供你目检 | 排版检查 |
| `labreport card list \| show <实验> \| new <实验>` | **实验知识卡**：本实验的公式、表结构、量级易错点、参考数据 | 动手前先查卡 |

退出码：`0` 就绪/无 error；`1` 有错；`2` 用法错误；`3` 缺可选程序（功能降级但可用）。

## 2. 标准流程

```
0) labreport doctor
   labreport card list / show <实验名>     → 有卡先读卡：公式、表结构、量级易错点、参考数据
1) labreport inspect 模板.docx          → 拿到物理网格：表几行几列、哪些是合并格
2) 读照片（放大、逐格确认，记录 source）→ 写 data.json
   labreport data check data.json       → 有 error 就回看照片，改到 0 error
3) 写 map.json：把每个值/公式落到 表N 行R 列C
4) labreport fill ... （先不加 --apply）
   → 把计划给用户看 → 确认后 labreport fill ... --apply
5) labreport audit 成品.docx --source 模板.docx --data data.json   → 必须 0 error
6) labreport pdf 成品.docx --pages      → 自己看图确认排版（中文、公式、表格不破版）
7) 没有卡时：做完用 `labreport card new <实验名>` 把本次的公式/表结构/易错点/参考数据落成新卡
```

## 3. data.json 最小示例（详见 `labreport data template`）

```jsonc
{
  "experiment": "扭摆法测量转动惯量",
  "quantities": {
    "m_cyl": { "value": 1116.12, "unit": "g",  "source": "img3 表3.4 第1行 质量" },
    "D_cyl": { "value": 100.50,  "unit": "mm", "source": "img3 表3.4 第1行 直径" }
  },
  "measurements": [
    { "id": "T0", "label": "空盘", "unit": "ms",
      "values": [774, 772, 773, 775, 774, 775], "mean": 774, "precision": 0, "source": "img5 第2行" }
  ],
  "targets": [
    { "name": "J1理", "expr": "1/8*m_cyl*D_cyl^2", "unit": "kg*m^2",
      "expected": 1.4091e-3, "tolerance": 0.01 }
  ]
}
```

要点：
- **不必自己换算单位**：`value + unit` 照仪表/照片原样填，表达式里用的已是 SI 值（`1116.12 g` → `1.11612`）。
- `source` 必填：校验失败时才知道该回看哪张照片的哪个位置。
- `targets` 按顺序求值，后面的可以引用前面的名字；`expected` 是你的核对基准。

## 4. map.json 最小示例

```jsonc
{ "targets": [
  { "table": 2, "row": 1, "col": 2, "set": "{{experiment}}" },
  { "table": 2, "row": 2, "col": 2, "set": "XX 班" },
  { "find": "成绩", "append": "（补充说明）" },
  { "table": 2, "row": 6, "col": 1, "append": "圆柱理论转动惯量：" },
  { "table": 2, "row": 6, "col": 1, "append": true,
    "formula": "J_{1\\mathrm{理}}=\\frac{1}{8}mD^{2}={{targets.J1理}}" }
]}
```

要点：
- `table/row/col` 是**屏幕上看到的**行列（1 起），合并单元格由 `fill` 自己换算到真实单元格；不确定就先 `inspect` 看。
- 也可用 `find` 按单元格文字定位（`occurrence` 指定第几次命中）。
- 动作四选一：`set`（替换）/ `append`（追加段落）/ `clear` / `formula`（插原生公式）。
- 占位符 `{{...}}` 可用键：`experiment`、`quantities.<key>`、`measurements.<id>`、`targets.<name>`（`--data` 提供）。

## 5. 出问题怎么查

| 现象 | 处置 |
|---|---|
| `data check` 报平均值不自洽 / 离群 | 回看该组 `source` 指的照片区域，逐格重认 |
| `data check` 报与预期值相差太大 | 检查公式、单位、抄录；确认无误再改 `expected` 或说明原因 |
| `fill` 报"超出网格" | `inspect` 看该表实际行列与合并情况，改用锚点 `find` |
| `audit` 报 LaTeX 残留 | 公式没转成功：检查 LaTeX 里是否有不支持的命令，或改用 `omml` 单独产出 |
| `audit` 报数据未命中 | 该值漏填，或写成了别的形式（如 `1.4091×10⁻³` vs `0.0014091`，工具两种都认） |
| `doctor` 报缺 LibreOffice/Word | 只影响 `pdf` 目检，不影响填报告 |
| `pdf` 退出码 3 | 环境里没有转换器（装 Word 或 LibreOffice）；**受限沙箱下 Word COM 也可能起不来**——需要放开文件权限，或改用 LibreOffice 头less 转换 |

## 6. 边界

- 只做"**按模板填**"与"**自检**"；没有模板、要新建报告时先请用户提供模板。
- 不替用户编造数据：识别不出的格子留占位并在回复里说明，不许猜一个数填进去。
- 用户没确认前不写文件（`fill` 默认行为已经保证这一点）。
