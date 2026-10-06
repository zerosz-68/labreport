"""按物理网格坐标填写报告：文本 / 追加 / 清空 / 插入原生公式。

为什么不用 python-docx 的下标：合并单元格会被展开成重复对象，`table.rows[r].cells[c]`
的下标与真实网格对不上。这里统一用 inspect 的"物理网格 -> 真实单元格 锚点"定位，
所以 `{"table":2,"row":1,"col":2}` 指的是**屏幕上看到的第 2 表第 1 行第 2 列**，
落进哪个真实单元格由合并关系决定。

确认机制：默认**只出计划不落地**（打印 每处 旧值→新值，供人或 agent 过目）；
加 `--apply`（或 `--auto`）才真正写文件。
"""

from __future__ import annotations

import copy
import json
import os
import re

from docx import Document
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn

from ..data.check import check as check_data
from ..omml.latex import latex_to_omml
from .inspect import W, M, table_structure

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

PLACEHOLDER = re.compile(r"\{\{\s*([A-Za-z0-9_.\u4e00-\u9fff]+)\s*\}\}")


# ---------------------------------------------------------------- 取值
def build_values(data: dict | None) -> dict:
    """把 data.json 摊平成 {{...}} 可引用的字典（复用 data check 的计算结果）。"""
    if not data:
        return {}
    report = check_data(data)
    vals = {"experiment": data.get("experiment", "")}
    for q in report["quantities"]:
        vals[f"quantities.{q['key']}"] = q["value"]
        vals[f"quantities.{q['key']}.SI"] = q["valueSI"]
    for m in report["measurements"]:
        vals[f"measurements.{m['id']}"] = m["mean"]
        vals[f"measurements.{m['id']}.mean"] = m["mean"]
        vals[f"measurements.{m['id']}.sd"] = m["sd"]
        vals[f"measurements.{m['id']}.SI"] = m["meanInBase"]
    for t in report["targets"]:
        vals[f"targets.{t['name']}"] = t["value"]
    return vals


def _fmt(value, fmt: str | None) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    if fmt:
        try:
            return fmt.format(value) if "{}" in fmt or "{:" in fmt else (fmt % value)
        except Exception:  # noqa: BLE001
            return str(value)
    if isinstance(value, float):
        return f"{value:g}"
    return str(value)


def render_tpl(text: str, values: dict, fmt: str | None = None) -> str:
    """替换 {{key}}；若整串就是一个占位符且给了 format，则按 format 输出。"""
    if not isinstance(text, str):
        return text

    def sub(match):
        key = match.group(1)
        if key not in values:
            raise KeyError(f"数据里没有 {key}（可用键：{', '.join(sorted(values)[:8])} …）")
        return _fmt(values[key], None)

    only = PLACEHOLDER.fullmatch(text.strip())
    if only and fmt:
        key = only.group(1)
        if key not in values:
            raise KeyError(f"数据里没有 {key}")
        return _fmt(values[key], fmt)
    return PLACEHOLDER.sub(sub, text)


# ---------------------------------------------------------------- 定位
def _top_tables(doc):
    return [ch for ch in doc.element.body if ch.tag == W + "tbl"]


def locate_cell(target: dict, tables):
    """返回 (tc, 位置描述)。支持坐标定位或按文字锚点定位。"""
    if "table" in target:
        idx = int(target["table"])
        if not (1 <= idx <= len(tables)):
            raise LookupError(f"模板没有第 {idx} 张表（共 {len(tables)} 张）")
        anchors = table_structure(tables[idx - 1])["anchors"]
        # anchors 的键是 "行,列" 字符串（JSON 友好），坐标是 0 起的物理网格
        key = f"{int(target['row']) - 1},{int(target['col']) - 1}"
        if key not in anchors:
            raise LookupError(
                f"表{idx} 的 r{target['row']} c{target['col']} 超出网格"
                f"（该表 {table_structure(tables[idx-1])['rows']} 行 × {table_structure(tables[idx-1])['gridCols']} 列）"
            )
        a = anchors[key]
        return a["tc"], f"表{idx} r{target['row']}c{target['col']}（真实单元格锚点 r{a['r']+1}c{a['c']+1}）"

    if "find" in target:
        pat = str(target["find"])
        want = int(target.get("occurrence", 1))
        seen = 0
        for ti, tbl in enumerate(tables, 1):
            if target.get("table") and ti != int(target["table"]):
                continue
            for _, a in table_structure(tbl)["anchors"].items():
                if pat and pat in (a["text"] or ""):
                    seen += 1
                    if seen == want:
                        return a["tc"], f"表{ti} r{a['r']+1}c{a['c']+1}（命中文字 {pat!r}）"
        raise LookupError(f"没有找到文字锚点 {pat!r}（第 {want} 次命中）")

    raise LookupError("target 必须给 table+row+col 或 find")


# ---------------------------------------------------------------- 写入
def _ensure_paragraph(tc, append: bool):
    ps = tc.findall(W + "p")
    if not ps or append:
        p = OxmlElement("w:p")
        tc.append(p)
        return p, None
    first = ps[0]
    rpr = None
    for r in first.findall(W + "r"):
        if rpr is None:
            found = r.find(W + "rPr")
            if found is not None:
                rpr = copy.deepcopy(found)
        first.remove(r)
    for el in list(first.findall(M + "oMath")):
        first.remove(el)
    return first, rpr


def set_cell_text(tc, text: str) -> None:
    """替换单元格首个段落的文字（保留段落属性与首个 run 的字体）。"""
    p, rpr = _ensure_paragraph(tc, append=False)
    r = OxmlElement("w:r")
    if rpr is not None:
        r.append(rpr)
    t = OxmlElement("w:t")
    t.set(XML_SPACE, "preserve")
    t.text = text
    r.append(t)
    p.append(r)


def append_cell_text(tc, text: str) -> None:
    p, rpr = _ensure_paragraph(tc, append=True)
    r = OxmlElement("w:r")
    if rpr is not None:
        r.append(rpr)
    t = OxmlElement("w:t")
    t.set(XML_SPACE, "preserve")
    t.text = text
    r.append(t)
    p.append(r)


def clear_cell(tc) -> None:
    p, _ = _ensure_paragraph(tc, append=False)


def write_formula(tc, latex: str, append: bool = False) -> None:
    om = parse_xml(latex_to_omml(latex))
    p, _ = _ensure_paragraph(tc, append=append)
    p.append(om)


# ---------------------------------------------------------------- 嵌入表格
def _force_full_width(tbl) -> None:
    """表格宽度撑满所在单元格（100%）。"""
    pr = tbl._tbl.tblPr
    for el in pr.findall(W + "tblW"):
        pr.remove(el)
    el = OxmlElement("w:tblW")
    el.set(qn("w:type"), "pct")
    el.set(qn("w:w"), "5000")
    pr.append(el)


def _force_borders(tbl) -> None:
    """显式给单线边框：不依赖模板里是否存在 Table Grid 样式。"""
    pr = tbl._tbl.tblPr
    for old in pr.findall(W + "tblBorders"):
        pr.remove(old)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), "4")
        e.set(qn("w:space"), "0")
        e.set(qn("w:color"), "auto")
        borders.append(e)
    pr.append(borders)


def _ensure_cell_ends_with_paragraph(tc) -> None:
    """OOXML 要求单元格以段落结尾，否则 Word 会提示修复文档。"""
    kids = list(tc)
    if not kids or kids[-1].tag != W + "p":
        tc.append(OxmlElement("w:p"))


def insert_table_spec(tc, spec: dict, doc, values: dict) -> dict:
    """在单元格末尾插入「表名一行 + 一张（可含合并的）表格」。

    spec 格式见 docx/tables.py：{name, cols, grid, merges}。
    """
    cols = int(spec.get("cols") or 0)
    grid = spec.get("grid") or []
    if cols <= 0:
        raise ValueError("insert_table 需要 cols（网格列数）")
    if not grid:
        raise ValueError("insert_table 需要 grid（每行单元格文字）")

    name = spec.get("name") or ""
    if name:
        append_cell_text(tc, render_tpl(str(name), values))

    rows = len(grid)
    tbl = doc.add_table(rows=rows, cols=cols)  # 先建在文档尾部，稍后整体搬进单元格
    style = spec.get("style") or "Table Grid"
    try:
        tbl.style = style
    except Exception:  # noqa: BLE001  模板里没有该样式时忽略，边框由 _force_borders 保证
        pass

    for r, row in enumerate(grid):
        cells = list(row) + [""] * max(0, cols - len(row))
        for c in range(cols):
            raw = cells[c]
            tbl.cell(r, c).text = render_tpl("" if raw is None else str(raw), values)

    # spec.formulas：{"行,列": "LaTeX"} —— 把这些格子写成 Word 原生公式
    # （原表里往往是"J₁=1/8mD²"这样的文字，要求公式格式正确时用它替换）
    # 必须放在合并**之后**：先合并时锚点格的内容会被重建，先写的公式会丢。
    for item in spec.get("merges") or []:
        vals = list(item) + [1, 1]
        r0, c0 = int(vals[0]) - 1, int(vals[1]) - 1
        rowspan, colspan = int(vals[2] or 1), int(vals[3] or 1)
        r1, c1 = r0 + rowspan - 1, c0 + colspan - 1
        if not (0 <= r0 <= r1 < rows and 0 <= c0 <= c1 < cols):
            raise ValueError(f"合并范围越界：{item}（该表 {rows} 行 × {cols} 列）")
        # 先赋文字再合并，可以保证合并后留下的是锚点格的文字
        a, b = tbl.cell(r0, c0), tbl.cell(r1, c1)
        if a._tc is not b._tc:
            a.merge(b)

    # spec.formulas：{"行,列": "LaTeX"} —— 把这些格子写成 Word 原生公式
    # （原表里往往是"J₁=1/8mD²"这样的文字，要求公式格式正确时用它替换）
    # 注意必须放在合并**之后**：合并会重建锚点格的内容，先写的公式会丢。
    for pos, latex in (spec.get("formulas") or {}).items():
        fr, fc = (int(x) - 1 for x in str(pos).replace("，", ",").split(","))
        if not (0 <= fr < rows and 0 <= fc < cols):
            raise ValueError(f"formulas 坐标越界：{pos}（该表 {rows} 行 × {cols} 列）")
        write_formula(
            tbl.cell(fr, fc)._tc, render_tpl(str(latex), values), append=False
        )

    tbl.autofit = True
    _force_full_width(tbl)
    _force_borders(tbl)
    tc.append(tbl._tbl)
    _ensure_cell_ends_with_paragraph(tc)
    return {"rows": rows, "cols": cols, "merges": len(spec.get("merges") or [])}


def load_table_specs(path: str) -> list:
    """读 `labreport tables` 产出的 JSON（也接受裸数组）。"""
    with open(path, encoding="utf-8") as fh:
        payload = json.load(fh)
    specs = payload if isinstance(payload, list) else (payload.get("tables") or [])
    for s in specs:
        s.pop("_源表序号", None)
    return specs


def _expand_tables(targets: list, base_dir: str) -> list:
    """把 `insert_tables_file` 就地展开成连续的 insert_table 动作。

    这样表格能插在 map 里指定的**位置**（例如"第五节标题之后、第六节之前"），
    而不是像 `--tables` 那样只能追加到末尾。
    """
    out = []
    for t in targets:
        ref = t.get("insert_tables_file")
        if not ref:
            out.append(t)
            continue
        path = ref if os.path.isabs(ref) else os.path.join(base_dir, ref)
        for spec in load_table_specs(path):
            out.append(
                {
                    "table": t.get("table"),
                    "row": t.get("row"),
                    "col": t.get("col"),
                    "insert_table": spec,
                }
            )
    return out


def parse_into(text: str) -> tuple:
    """解析 --into "表,行,列"（1 起，物理网格坐标）。"""
    parts = [p.strip() for p in str(text).replace("，", ",").split(",")]
    if len(parts) != 3 or not all(p.isdigit() for p in parts):
        raise ValueError('--into 需要“表,行,列”三个数字，例如 --into "2,6,1"')
    return tuple(int(p) for p in parts)


# ---------------------------------------------------------------- 主流程
def _action_of(target: dict):
    if "insert_table" in target:
        return "table"
    if "formula" in target:
        return "formula"
    if "set" in target:
        return "set"
    if "append" in target:
        return "append"
    if "clear" in target:
        return "clear"
    raise ValueError("target 必须含 set / append / clear / formula / insert_table 之一")


def _cell_text(tc) -> str:
    from .inspect import _text_of

    return _text_of(tc)


def run(args) -> int:
    tpl_path = args.template
    if not os.path.isfile(tpl_path):
        print(f"找不到模板：{tpl_path}")
        return 1
    if not os.path.isfile(args.map):
        print(f"找不到映射文件：{args.map}")
        return 1

    with open(args.map, encoding="utf-8") as fh:
        mapping = json.load(fh)
    targets = _expand_tables(
        list(mapping.get("targets") or []), os.path.dirname(os.path.abspath(args.map))
    )

    # --tables：把「从课堂报告书提取出来的表格」按顺序一次性嵌入某个单元格
    # （表名 + 表格由 insert_table_spec 负责，仍是"先出计划、确认后落地"）
    if getattr(args, "tables", None):
        if not os.path.isfile(args.tables):
            print(f"找不到表格规格文件：{args.tables}")
            return 1
        try:
            into = parse_into(getattr(args, "into", "") or "")
        except ValueError as exc:
            print(str(exc))
            return 1
        specs = load_table_specs(args.tables)
        if not specs:
            print(f"表格规格文件里没有表格：{args.tables}")
            return 1
        targets += [
            {"table": into[0], "row": into[1], "col": into[2], "insert_table": s} for s in specs
        ]

    if not targets:
        print("映射文件里没有 targets（也没有用 --tables 提供表格）")
        return 1

    data = None
    data_path = args.data or mapping.get("data")
    if data_path:
        if not os.path.isabs(data_path):
            data_path = os.path.join(os.path.dirname(os.path.abspath(args.map)), data_path)
        if os.path.isfile(data_path):
            with open(data_path, encoding="utf-8") as fh:
                data = json.load(fh)
        else:
            print(f"警告：找不到 data 文件 {data_path}，占位符将无法解析")

    values = build_values(data)
    doc = Document(tpl_path)
    tables = _top_tables(doc)

    plan, errors = [], 0
    for i, target in enumerate(targets, 1):
        entry = {"n": i, "action": "", "where": "", "before": "", "after": "", "status": "ok"}
        try:
            action = _action_of(target)
            entry["action"] = action
            tc, where = locate_cell(target, tables)
            entry["where"] = where
            entry["before"] = _cell_text(tc)[:60]
            entry["hasNested"] = bool(tc.findall(W + "tbl")) or len(list(tc.iter(W + "tbl"))) > 0

            if action == "set":
                text = render_tpl(target["set"], values, target.get("format"))
                entry["after"] = text[:60]
                if apply_mode(args):
                    set_cell_text(tc, text)
            elif action == "append":
                text = render_tpl(target["append"], values, target.get("format"))
                entry["after"] = ("＋" + text)[:60]
                if apply_mode(args):
                    append_cell_text(tc, text)
            elif action == "clear":
                entry["after"] = "（清空）"
                if apply_mode(args):
                    clear_cell(tc)
            elif action == "formula":
                latex = render_tpl(target["formula"], values, target.get("format"))
                entry["after"] = ("公式 " + latex)[:140]
                if apply_mode(args):
                    write_formula(tc, latex, append=bool(target.get("append")))
            elif action == "table":
                spec = dict(target["insert_table"])
                spec["name"] = render_tpl(str(spec.get("name") or ""), values)
                entry["after"] = (
                    f"嵌入表格：{spec['name']}"
                    f"（{len(spec.get('grid') or [])} 行 × {int(spec.get('cols') or 0)} 列，"
                    f"合并 {len(spec.get('merges') or [])} 处）"
                )[:140]
                if apply_mode(args):
                    insert_table_spec(tc, spec, doc, values)
        except Exception as exc:  # noqa: BLE001
            entry["status"] = "error"
            entry["after"] = str(exc)
            errors += 1
        plan.append(entry)

    out_path = args.out or os.path.join(
        os.path.dirname(os.path.abspath(tpl_path)),
        os.path.splitext(os.path.basename(tpl_path))[0] + "-filled.docx",
    )

    print(f"模板   : {tpl_path}")
    print(f"映射   : {args.map}" + (f"    数据: {data_path}" if data_path else "    （无数据文件）"))
    print(f"可用键 : {len(values)} 个" + (f"  例: {', '.join(sorted(values)[:6])}" if values else ""))
    print(f"模式   : {'写入（--apply/--auto）' if apply_mode(args) else '仅出计划（默认，未写文件）'}")
    print("")
    for e in plan:
        flag = {"ok": "✓", "error": "✗"}.get(e["status"], "?")
        nested = "  〔该格含嵌套表，set 只改本层文字〕" if e.get("hasNested") and e["action"] == "set" else ""
        print(f"  {flag} #{e['n']:<2} [{e['action']:<7}] {e['where']}{nested}")
        print(f"        旧: {e['before'] or '（空）'}")
        print(f"        新: {e['after'] or '（空）'}")

    plan_path = args.plan or (
        os.path.splitext(out_path)[0] + ".plan.json" if apply_mode(args) else ""
    )
    if plan_path:
        with open(plan_path, "w", encoding="utf-8") as fh:
            json.dump({"template": tpl_path, "map": args.map, "plan": plan}, fh, ensure_ascii=False, indent=1)
        print(f"\n计划已写出：{plan_path}")

    if apply_mode(args):
        if errors:
            print(f"\n有 {errors} 处定位/取值失败，未写出文件。请先修映射或数据。")
            return 1
        doc.save(out_path)
        print(f"\n已写出：{out_path}")
        print("提示：接着跑 `labreport inspect` 或 `labreport audit` 检查填充结果。")
    else:
        print("\n（未写文件）确认无误后加 --apply 落地；全自动场景用 --auto。")
    return 0 if errors == 0 else 1


def apply_mode(args) -> bool:
    return bool(getattr(args, "apply", False) or getattr(args, "auto", False))
