"""读 Word 模板结构 —— 重点是把"合并单元格"还原成真实的物理网格。

为什么必须自己做：python-docx 的 `table.rows[x].cells[y]` 会把合并区展开成重复
对象，下标和真实网格对不上，按它填数据必然错位。这里直接从 OOXML 的
`w:tcPr/w:gridSpan` 与 `w:tcPr/w:vMerge` 计算锚点，得到"物理行列 -> 真实单元格"。

另一个坑：单元格里常常嵌着**嵌套表格**（学生把整段正文和表格粘在合并单元格里）。
统计时必须跳过嵌套表，否则父表会把子表的内容/公式重复算进来。
"""

from __future__ import annotations

import json
import os
import re

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"

# 占位/待填的典型形态（用于识别"填写边界"）
PLACEHOLDER_PATTERNS = [
    re.compile(r"^[\s\u3000]*$"),
    re.compile(r"^[—\-–_…\.·\s]{1,}$"),
    re.compile(r"^[（(]\s*[)）]$"),
    re.compile(r"^[【\[［]\s*[】\]］]$"),
]


def _iter_own(el, tag):
    """遍历 el 的后代，但跳过嵌套表格 w:tbl，避免父表重复统计子表内容。"""
    for child in el:
        if child.tag == W + "tbl":
            continue
        if child.tag == tag:
            yield child
        yield from _iter_own(child, tag)


def _text_of(el) -> str:
    return "".join(t.text or "" for t in _iter_own(el, W + "t")).strip()


def _math_texts(el) -> list:
    return [
        "".join(t.text or "" for t in _iter_own(om, M + "t"))
        for om in _iter_own(el, M + "oMath")
    ]


def _cell_info(tc) -> dict:
    tc_pr = tc.find(W + "tcPr")
    span, vmerge = 1, None
    if tc_pr is not None:
        gs = tc_pr.find(W + "gridSpan")
        if gs is not None:
            try:
                span = int(gs.get(W + "val", "1"))
            except ValueError:
                span = 1
        vm = tc_pr.find(W + "vMerge")
        if vm is not None:
            vmerge = vm.get(W + "val") or "continue"
    return {
        "text": _text_of(tc),
        "gridSpan": span,
        "vMerge": vmerge,
        "math": _math_texts(tc),
        "nImages": len(list(_iter_own(tc, W + "drawing")))
        + len(list(_iter_own(tc, W + "pict"))),
        "nNested": len(list(tc.iter(W + "tbl"))),
    }


def _grid_cols(tbl) -> int:
    grid = tbl.find(W + "tblGrid")
    if grid is None:
        return 0
    return len(grid.findall(W + "gridCol"))


def _is_blank(a: dict) -> bool:
    return not a["text"] and not a["math"] and a["nImages"] == 0 and a["nNested"] == 0


def table_structure(tbl) -> dict:
    """返回一张表的物理结构：行列数、真实单元格、合并、空白、公式、图片、嵌套表。"""
    ncols = _grid_cols(tbl)
    rows_cells = []
    for tr in tbl.findall(W + "tr"):
        rows_cells.append([(tc, _cell_info(tc)) for tc in tr.findall(W + "tc")])

    if ncols <= 0:
        ncols = max((sum(cc["gridSpan"] for _, cc in r) for r in rows_cells), default=0)

    # 逐格铺进物理网格，并记录每个物理格的"锚点"（真实单元格左上角）
    anchor_of = {}
    anchors = {}
    for r, cells in enumerate(rows_cells):
        c = 0
        for tc, cell in cells:
            while c < ncols and (r, c) in anchor_of:
                c += 1
            span = cell["gridSpan"]
            if cell["vMerge"] == "continue" and (r - 1, c) in anchor_of:
                anchor = anchor_of[(r - 1, c)]
                anchors[anchor]["vContinues"] += 1
            else:
                anchor = (r, c)
                anchors[anchor] = {
                    "r": r,
                    "c": c,
                    "span": span,
                    "text": cell["text"],
                    "math": cell["math"],
                    "nImages": cell["nImages"],
                    "nNested": cell["nNested"],
                    "vMerge": cell["vMerge"],
                    "vContinues": 0,
                    "tc": tc,
                }
            for k in range(span):
                if c + k < ncols:
                    anchor_of[(r, c + k)] = anchor
            c += span

    real = list(anchors.values())
    merged = [a for a in real if a["span"] > 1 or a["vMerge"] == "restart" or a["vContinues"]]
    blank = [a for a in real if _is_blank(a)]
    placeholder = [
        a for a in real if a["text"] and any(p.match(a["text"]) for p in PLACEHOLDER_PATTERNS)
    ]

    return {
        "gridCols": ncols,
        "rows": len(rows_cells),
        "logicalCells": sum(len(r) for r in rows_cells),
        "realCells": len(real),
        "mergedCells": len(merged),
        "blankCells": len(blank),
        "placeholderCells": len(placeholder),
        "mathInTable": sum(len(a["math"]) for a in real),
        "imageCells": sum(1 for a in real if a["nImages"]),
        "nestedTables": sum(a["nNested"] for a in real),
        "cells": real,
        # 键是 "行,列" 字符串（JSON 友好，0 起）；值是同一个单元格信息字典（含 tc 元素，
        # 供 fill 直接改文档；写 JSON 前由 inspect_document 去掉 tc）
        "anchors": {
            f"{pos[0]},{pos[1]}": anchors[tuple(anchor)] for pos, anchor in anchor_of.items()
        },
    }


def inspect_document(path: str) -> dict:
    from docx import Document

    doc = Document(path)
    body = doc.element.body

    paras = [{"text": p.text.strip(), "style": p.style.name if p.style else None} for p in doc.paragraphs]
    non_empty = [p for p in paras if p["text"]]

    # 只取文档流里的顶层表格（嵌套表交给所属单元格统计）
    top_tables = [ch for ch in body if ch.tag == W + "tbl"]
    tables = [table_structure(tbl) for tbl in top_tables]
    nested_total = sum(t["nestedTables"] for t in tables)

    # 元素引用只给 fill 用，不进 JSON/文本报告
    for t in tables:
        for a in t["cells"]:
            a.pop("tc", None)
        for a in t["anchors"].values():
            a.pop("tc", None)

    math_all = list(body.iter(M + "oMath"))
    images = len(list(body.iter(W + "drawing"))) + len(list(body.iter(W + "pict")))

    placeholder_paras = [
        p["text"] for p in non_empty if any(x.match(p["text"]) for x in PLACEHOLDER_PATTERNS)
    ]

    return {
        "file": os.path.abspath(path),
        "sizeKB": round(os.path.getsize(path) / 1024, 1),
        "paragraphs": {"total": len(paras), "nonEmpty": len(non_empty)},
        "sampleParagraphs": [p["text"] for p in non_empty[:10]],
        "placeholderParagraphs": placeholder_paras,
        "tables": tables,
        "tableCount": len(tables),
        "nestedTableCount": nested_total,
        "mathCount": len(math_all),
        "mathSamples": ["".join(t.text or "" for t in om.iter(M + "t")) for om in math_all],
        "images": images,
    }


def _truncate(text: str, limit: int) -> str:
    text = text.replace("\n", " ").strip()
    return text if len(text) <= limit else text[: limit - 1] + "…"


def render(report: dict, max_text: int, formula_sample: int) -> str:
    out = []
    out.append(f"文件      : {report['file']}")
    out.append(f"大小      : {report['sizeKB']} KB")
    out.append(
        f"段落      : {report['paragraphs']['total']} 段（非空 {report['paragraphs']['nonEmpty']}）"
        f"   图片 {report['images']}   公式 {report['mathCount']} 个(OMML)"
    )
    if report["placeholderParagraphs"]:
        out.append(
            f"占位段落  : {len(report['placeholderParagraphs'])} 处 -> {report['placeholderParagraphs'][:5]}"
        )
    out.append("")
    out.append(
        f"表格      : 顶层 {report['tableCount']} 张（另有嵌套表 {report['nestedTableCount']} 张，计入所属单元格）"
    )

    for i, t in enumerate(report["tables"], 1):
        out.append("")
        out.append(
            f"  ── 表{i}: {t['rows']} 行 × {t['gridCols']} 列 | 逻辑单元格 {t['logicalCells']}"
            f" → 真实单元格 {t['realCells']} | 合并 {t['mergedCells']} | 空白 {t['blankCells']}"
            f" | 占位 {t['placeholderCells']} | 本表公式 {t['mathInTable']}"
            f" | 含图 {t['imageCells']} | 嵌套表 {t['nestedTables']}"
        )
        by_row = {}
        for a in t["cells"]:
            by_row.setdefault(a["r"], []).append(a)
        for r in sorted(by_row):
            parts = []
            for a in sorted(by_row[r], key=lambda x: x["c"]):
                tag = ""
                if a["span"] > 1:
                    tag += f"⇢{a['span']}列"
                if a["vMerge"] == "restart":
                    tag += "⇣合并"
                if a["nNested"]:
                    body_text = f"〔嵌套表×{a['nNested']}〕" + _truncate(a["text"], max_text)
                elif _is_blank(a):
                    body_text = "〔空〕"
                elif a["math"] and not a["text"]:
                    body_text = "〔公式〕"
                else:
                    body_text = _truncate(a["text"], max_text)
                parts.append(f"{body_text}{tag}")
            out.append(f"    r{r + 1:<2} | " + " | ".join(parts))

    if report["mathSamples"]:
        out.append("")
        out.append(f"公式样例（前 {min(formula_sample, len(report['mathSamples']))} 条）:")
        for s in report["mathSamples"][:formula_sample]:
            out.append(f"    {_truncate(s, 90)}")

    return "\n".join(out)


def run(args) -> int:
    path = args.docx
    if not os.path.isfile(path):
        print(f"找不到文件：{path}")
        return 1
    if not path.lower().endswith(".docx"):
        print(f"只支持 .docx（收到 {os.path.splitext(path)[1]}）")
        return 1
    try:
        report = inspect_document(path)
    except Exception as exc:  # noqa: BLE001
        print(f"读取失败：{exc}")
        return 1

    if getattr(args, "json", False):
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(render(report, args.max_text, args.formula_sample))
    return 0
