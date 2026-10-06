"""从一份 docx 里提取表格，转成可直接嵌入报告的规格（JSON）。

为什么需要它：课堂报告书里的表格要"完全照原表结构和合并关系"搬进实验报告正文，
让人或 agent 手抄 JSON 既慢又容易抄错合并关系。直接从原表读出来最可靠。

规格格式（也是 `fill` 的 `insert_table` 动作接受的格式）：

    {
      "name": "1.待测物体质量和尺寸的测量",     // 表格名称，会写在表格上方单独一行
      "cols": 6,                                // 网格列数
      "grid": [["试样", "质量m（g）", "试样的尺寸（mm）", "", "J理（10⁻³kg·m²）", ""], ...],
      "merges": [[2, 3, 1, 2]]                  // [行, 列, 跨行, 跨列]，1 起；被覆盖的格在 grid 里留空
    }

两个约定：

- **表名提取**：很多报告书把表名放在表格第一行（整行合并）。默认把这行"提上来"当作
  `name`（这样嵌入后就是"表名一行 + 表格"），用 `--keep-title-row` 可以保留在表内。
- **空表跳过**：整张表只有表头没有数据时默认跳过（`--keep-empty` 可保留），
  因为空白表搬进报告没有意义。

公式：原表里的公式往往只是文字（如 `J₁=1/8mD²`）。要让它变成 Word 原生公式，
在规格里加 `"formulas": {"行,列": "LaTeX"}`（行列相对该表，1 起），`fill` 会把这些
格子写成 OMML。例如：

    {"name": "1.待测物体质量和尺寸的测量", "cols": 6, "grid": [...],
     "merges": [[2, 3, 1, 2]],
     "formulas": {"2,5": "J_{1}=\\frac{1}{8}mD^{2}"}}
"""

from __future__ import annotations

import json
import os

from .inspect import W, table_structure

TITLE_MAX = 40


def _preceding_text(body, tbl) -> str:
    """表格前面最近的一个非空段落（有些文档把表名放在表格外面）。"""
    prev = tbl.getprevious()
    while prev is not None and prev.tag == W + "tbl":
        prev = prev.getprevious()
    if prev is None:
        return ""
    if prev.tag != W + "p":
        return ""
    text = "".join(t.text or "" for t in prev.iter(W + "t")).strip()
    return text if len(text) <= TITLE_MAX else ""


def extract_table_spec(tbl, lift_title: bool = True, fallback_name: str = "") -> dict:
    """把一张表转成嵌入规格。"""
    st = table_structure(tbl)
    rows, cols = st["rows"], st["gridCols"]
    grid = [["" for _ in range(cols)] for _ in range(rows)]
    merges = []
    for a in st["cells"]:
        r, c = a["r"], a["c"]
        grid[r][c] = a["text"]
        rowspan = 1 + int(a.get("vContinues") or 0)
        colspan = int(a.get("span") or 1)
        if rowspan > 1 or colspan > 1:
            merges.append([r + 1, c + 1, rowspan, colspan])

    name = fallback_name
    # 第一行是否是"整行合并的表名"
    first_row_anchors = [a for a in st["cells"] if a["r"] == 0]
    if lift_title and len(first_row_anchors) == 1 and int(first_row_anchors[0]["span"]) >= cols:
        text = (first_row_anchors[0]["text"] or "").strip()
        if text and len(text) <= TITLE_MAX:
            name = text
            grid = grid[1:]
            merges = [[r - 1, c, rs, cs] for (r, c, rs, cs) in merges if r > 1]
            rows -= 1

    return {"name": name, "cols": cols, "grid": grid, "merges": merges}


def _is_empty(grid) -> bool:
    return all(not (cell or "").strip() for row in grid for cell in row)


def extract_document_tables(
    path: str, lift_title: bool = True, keep_empty: bool = False
) -> dict:
    """提取一份文档里的所有顶层表格。"""
    from docx import Document

    doc = Document(path)
    body = doc.element.body
    specs, skipped = [], []
    for i, tbl in enumerate([ch for ch in body if ch.tag == W + "tbl"], 1):
        spec = extract_table_spec(tbl, lift_title=lift_title)
        if not spec["name"]:
            spec["name"] = _preceding_text(body, tbl)
        if not spec["name"]:
            spec["name"] = f"表{i}"
        if _is_empty(spec["grid"]) and not keep_empty:
            skipped.append({"index": i, "name": spec["name"], "rows": len(spec["grid"])})
            continue
        spec["_源表序号"] = i
        specs.append(spec)
    return {
        "source": os.path.abspath(path),
        "count": len(specs),
        "skipped": skipped,
        "tables": specs,
    }


def run(args) -> int:
    path = args.source
    if not os.path.isfile(path):
        print(f"找不到文件：{path}")
        return 1
    report = extract_document_tables(
        path, lift_title=not getattr(args, "keep_title_row", False),
        keep_empty=getattr(args, "keep_empty", False),
    )

    print(f"源文件 : {report['source']}")
    print(f"提取到 : {report['count']} 张表" + (f"（跳过空表 {len(report['skipped'])} 张）" if report["skipped"] else ""))
    print("")
    for spec in report["tables"]:
        filled = sum(1 for row in spec["grid"] for cell in row if (cell or "").strip())
        total = len(spec["grid"]) * spec["cols"]
        print(f"  表{spec['_源表序号']}：{spec['name']}")
        print(f"        {len(spec['grid'])} 行 × {spec['cols']} 列 | 有内容 {filled}/{total} 格 | 合并 {len(spec['merges'])} 处")
    for sk in report["skipped"]:
        print(f"  （跳过 表{sk['index']}：{sk['name']} —— 整表无数据）")

    if args.out:
        payload = {k: v for k, v in report.items() if k != "skipped"}
        with open(args.out, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, ensure_ascii=False, indent=1)
        print(f"\n已写出：{args.out}")
        print("嵌入到报告：labreport fill <模板.docx> --map map.json "
              f"--tables {os.path.basename(args.out)} --into \"<表>,<行>,<列>\" --apply")
    elif getattr(args, "json", False):
        print(json.dumps(report, ensure_ascii=False, indent=1))
    return 0
