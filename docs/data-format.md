# 数据格式

两个 JSON 文件：**`data.json`**（实验数据，agent 读照片后填）与 **`map.json`**（填写映射，人和 agent 一起写）。

生成骨架：`labreport data template -o data.json`。

---

## data.json

```jsonc
{
  "experiment": "扭摆法测量转动惯量",
  "source": { "images": ["img1.jpg", "img3.jpg"], "note": "手写记录照片" },

  "quantities": {
    "m_cyl": { "value": 1116.12, "unit": "g",  "uncertainty": 0.01, "source": "img3 表3.4 第1行 质量" },
    "D_cyl": { "value": 100.50,  "unit": "mm", "uncertainty": 0.02, "source": "img3 表3.4 第1行 直径" }
  },

  "measurements": [
    { "id": "T0", "label": "空盘", "unit": "ms",
      "values": [774, 772, 773, 775, 774, 775], "mean": 774, "precision": 0,
      "source": "img5 表3.5 第2行" }
  ],

  "targets": [
    { "name": "J1理", "expr": "1/8*m_cyl*D_cyl^2", "unit": "kg*m^2",
      "expected": 1.4091e-3, "tolerance": 0.01 }
  ]
}
```

### 顶层字段

| 字段 | 必填 | 说明 |
|---|---|---|
| `experiment` | 推荐 | 实验名，可用 `{{experiment}}` 填进报告 |
| `source` | 可选 | 照片清单与说明，方便追溯 |
| `quantities` | 可选 | 单值测量（质量、直径、长度…） |
| `measurements` | 可选 | 多次测量（周期、时间…） |
| `targets` | 可选 | 需要重算并核对的目标量（转动惯量、误差、斜率…） |

### quantities（单值量）

| 字段 | 说明 |
|---|---|
| `value` | **按记录原样填**（不要自己换算） |
| `unit` | 记录单位，见下表；未知单位会给 warn |
| `uncertainty` | 可选，B 类不确定度（留给报告用） |
| `symbol` | 可选，额外注册一个别名（如 `L`） |
| `source` | 强烈建议填：哪张照片、哪个位置 |

支持的单位（换算到 SI）：

| 类别 | 单位 |
|---|---|
| 质量 | `mg` `g` `kg` |
| 长度 | `mm` `cm` `m` |
| 时间 | `ms` `s` `min` |
| 转动惯量 | `kg*m^2` `g*cm^2` |
| 其它 | `N` `N*m` `N*m/rad` `J` `Hz` `V` `A` `Ω` `°C` `%` `1`（空） |

> 缺单位时若确实需要，可在 `src/labreport/data/check.py` 的 `UNITS` 表里补一行。

### measurements（多次测量）

| 字段 | 说明 |
|---|---|
| `id` | 表达式里引用的名字（如 `T0`） |
| `label` | 人类可读说明（如"空盘"） |
| `unit` | 记录单位（`ms`/`s`…） |
| `values` | 每次测得值，数组 |
| `mean` | 可选，你在纸上写的平均值（用来核对你抄得对不对） |
| `precision` | 可选，末位精度（默认从数值推断），决定平均值自洽的容差 |
| `scale` | 可选，默认 1。**测 50 个周期的总时间时填 `0.02`**（=1/50） |
| `source` | 建议填 |

表达式里 `T0` 的值 = `mean × unit换算系数 × scale`，即**已经是 SI 单周期的值**。

### targets（重算目标）

| 字段 | 说明 |
|---|---|
| `name` | 结果名，可被后面的目标引用；也可用 `{{targets.<name>}}` 填进报告 |
| `expr` | 用**量的名字**写的算式。支持 `+ - * / ** %`、`^` 作幂、括号，以及 `sqrt` `abs` `sin` `cos` `tan` `log` `exp` `pi` `e` |
| `unit` | 结果单位（只记录，不做量纲推导） |
| `expected` | 可选，你的核对基准（理论值或手算值） |
| `tolerance` | 可选，默认 `0.01`（1%）；相对差超过它 → warn，超过 5 倍 → error |

按**数组顺序**求值，后面的可以引用前面的：

```jsonc
"targets": [
  { "name": "J1理", "expr": "1/8*m_cyl*D_cyl^2", "unit": "kg*m^2", "expected": 1.4091e-3 },
  { "name": "K",    "expr": "4*pi^2*J1理/(T1^2-T0^2)", "unit": "N*m/rad", "expected": 0.0348, "tolerance": 0.02 }
]
```

### 校验结果怎么读

```console
$ labreport data check torsion.json
单位归一化（表达式里用的 SI 值）:
  m_cyl      1116.12 g  ->  1.11612
  D_cyl      100.5 mm  ->  0.1005
测量组统计:
  T0     空盘               n=6  平均=773.8ms  = 0.773833  标准差=1.169  区间[772,775]  平均值自洽 ✓
重算目标:
  J1理        = 0.00140914 kg*m^2  预期=0.0014091  相对差=0.003%  [一致]
  [warn ] measurements.T2: 平均值不自洽：写了 1556，按 6 次算是 1573.33（容差 0.5）
结论: 可以填报告（error 0 / warn 0 / info 0）
```

`--json` 会给出同样内容的机器可读版本（含每项算出的数值），方便脚本判断。

---

## map.json

```jsonc
{
  "_说明": "table/row/col 是屏幕上看到的物理网格坐标（1 起）",
  "data": "torsion.json",
  "targets": [
    { "table": 2, "row": 1, "col": 2, "set": "{{experiment}}" },
    { "table": 2, "row": 2, "col": 5, "set": "2025000000" },
    { "find": "实验报告应包括以下内容", "append": "（补充说明）" },
    { "table": 2, "row": 6, "col": 1, "append": "圆柱理论转动惯量：" },
    { "table": 2, "row": 6, "col": 1, "append": true,
      "formula": "J_{1\\mathrm{理}}=\\frac{1}{8}mD^{2}={{targets.J1理}}" }
  ]
}
```

### 定位（二选一）

| 方式 | 写法 | 说明 |
|---|---|---|
| 物理坐标 | `{"table":2,"row":1,"col":2}` | **屏幕上看到的**行列，1 起；落进哪个真实单元格由合并关系决定。不确定就先 `inspect` |
| 文字锚点 | `{"find":"成绩","occurrence":1}` | 在单元格文字里找；`occurrence` 指定第几次命中；可加 `table` 限定表 |

### 动作（四选一）

| 动作 | 写法 | 行为 |
|---|---|---|
| 替换 | `{"set": "文本"}` | 替换该单元格首个段落的文字，**保留段落属性与首个 run 的字体** |
| 追加 | `{"append": "文本"}` | 在单元格里追加一个段落 |
| 清空 | `{"clear": true}` | 清掉该单元格的文字 |
| 公式 | `{"formula": "LaTeX"}` | 插入 Word 原生公式；加 `"append": true` 则新增段落而不是替换 |

### 占位符

`{{...}}` 可用的键（需要 `--data` 或 map.json 里的 `data` 提供）：

| 键 | 含义 |
|---|---|
| `experiment` | 实验名 |
| `quantities.<key>` | 单值量的**原值**（如 `1116.12`） |
| `quantities.<key>.SI` | 归一化后的值（如 `1.11612`） |
| `measurements.<id>` / `.mean` | 该组平均值（原单位） |
| `measurements.<id>.SI` | 换算成 SI 单次值（含 `scale`） |
| `measurements.<id>.sd` | 该组标准差 |
| `targets.<name>` | 重算目标的值 |

格式化：整串就是**一个**占位符时可用 `"format"`（如 `"{:.4g}"` 或 `"%.3f"`）：
`{"set": "{{targets.K}}", "format": "{:.4f}"}`。

### 定位失败会怎样

`fill` 会把每一处失败标成 `✗` 并给出原因（"超出网格"、"没找到文字锚点"…），**并且拒绝写文件**——映射没写对就不会产出半成品。
