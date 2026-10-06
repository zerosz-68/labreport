"""在报告末尾追加附件（如《课堂任务报告书》手写原件）。

需求来自实际流程：报告正文写完后，要把课堂报告书 1～2 页原件贴在最末尾，
上面标一行"附件"。要求：

- 标签单独一行、**不分页**（紧接正文之后）；
- 每张图**独占一页**（图与图之间插入分页符）；
- **满宽**：宽度撑满页边距内的可用宽度；若按宽度缩放后高度超出可用高度，
  则改为按高度自适应（避免图片被切到下一页）。

和其它命令一致：默认只出计划，`--apply`/`--auto` 才写文件。
"""

from __future__ import annotations

import os

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.shared import Emu
from docx.text.paragraph import Paragraph

EMU_PER_CM = 360000
CELL_MARGIN_CM = 0.4          # 单元格内边距，进框贴图时从可用宽度里扣掉


def usable_box(doc):
    """返回版心尺寸（EMU）：(宽, 高)。"""
    sec = doc.sections[-1]
    w = sec.page_width - sec.left_margin - sec.right_margin
    h = sec.page_height - sec.top_margin - sec.bottom_margin
    return int(w), int(h)


def fit_size(path: str, box_w: int, box_h: int) -> tuple:
    """按"满宽优先、超高则按高"计算插入尺寸（EMU）。"""
    from PIL import Image

    with Image.open(path) as im:
        px_w, px_h = im.size
    if px_w <= 0 or px_h <= 0:
        return box_w, box_h
    ratio = px_h / px_w
    w = box_w
    h = int(w * ratio)
    if h > box_h:                       # 太高 -> 按高度缩
        h = box_h
        w = int(h / ratio)
    return w, h


def _trim_trailing_empty_paragraphs(doc) -> int:
    """删掉文末多余的空段落。

    模板在正文表格后面自带两个空段落，正文填满后会溢出一张几乎空白的页；
    附件加进去以后更明显。这里只删"末尾连续的、完全没有内容的 w:p"。
    """
    from .inspect import M, W

    body = doc.element.body
    removed = 0
    for ch in reversed(list(body)):
        if ch.tag == W + "sectPr":
            continue
        if ch.tag != W + "p":
            break
        text = "".join(t.text or "" for t in ch.iter(W + "t")).strip()
        has_obj = (
            ch.find(".//" + W + "drawing") is not None
            or ch.find(".//" + W + "pict") is not None
            or ch.find(".//" + M + "oMath") is not None
        )
        if text or has_obj:
            break
        body.remove(ch)
        removed += 1
    return removed


def _new_paragraph(container, doc) -> Paragraph:
    """在文档末尾或某个单元格里新建段落。"""
    if container is None:
        return doc.add_paragraph()
    el = OxmlElement("w:p")
    container.append(el)
    return Paragraph(el, doc)


def resolve_cell(doc, into: str):
    """按物理网格坐标 "表,行,列" 定位单元格（复用 fill 的定位逻辑）。"""
    from .fill import _top_tables, locate_cell

    parts = [p.strip() for p in str(into).replace("，", ",").split(",")]
    if len(parts) != 3 or not all(p.isdigit() for p in parts):
        raise ValueError('--into 需要“表,行,列”三个数字，例如 --into "2,6,1"')
    t, r, c = (int(p) for p in parts)
    return locate_cell({"table": t, "row": r, "col": c}, _top_tables(doc))


def add_attachment(doc, images, label: str = "附件", page_break_between: bool = True,
                   align_center: bool = True, container=None) -> dict:
    """追加「标签行 + 每页一张的图片」。

    container 给 None 时追加到文档末尾；给某个 `w:tc` 元素时追加进该单元格
    （这样附件就在报告的表格框里，而不是漂在框外）。
    """
    box_w, box_h = usable_box(doc)
    if container is not None:
        box_w = max(int(box_w - CELL_MARGIN_CM * EMU_PER_CM), int(box_w * 0.6))
        # 单元格内还要放标签行与单元格内边距，按 80% 高度留余量：
        # 图若正好占满一页，表格底框会被挤到下一页，多出一张空白页
        box_h = int(box_h * 0.80)
    added = []

    if label:
        _new_paragraph(container, doc).add_run(label)

    first = True
    for path in images:
        w, h = fit_size(path, box_w, box_h)
        p = _new_paragraph(container, doc)
        if align_center:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if not first and page_break_between:
            p.add_run().add_break(WD_BREAK.PAGE)
        p.add_run().add_picture(path, width=Emu(w))
        added.append({"file": path, "widthEMU": w, "heightEMU": h, "pageBreakBefore": not first})
        first = False

    return {"label": label, "images": added, "boxW": box_w, "boxH": box_h}


def cm(emu: int) -> float:
    return round(emu / EMU_PER_CM, 2)


def run(args) -> int:
    doc_path = args.docx
    if not os.path.isfile(doc_path):
        print(f"找不到文件：{doc_path}")
        return 1
    images = [p for p in (args.images or []) if os.path.isfile(p)]
    missing = [p for p in (args.images or []) if not os.path.isfile(p)]
    for p in missing:
        print(f"找不到图片：{p}")
    if not images:
        print("没有可用的图片。用法：labreport attach 成品.docx 图片1 图片2 …")
        return 1

    doc = Document(doc_path)
    box_w, box_h = usable_box(doc)
    container, where = None, "文档末尾"
    if getattr(args, "into", None):
        try:
            container, where = resolve_cell(doc, args.into)
        except Exception as exc:  # noqa: BLE001
            print(f"定位失败：{exc}")
            return 1
        box_w = max(int(box_w - CELL_MARGIN_CM * EMU_PER_CM), int(box_w * 0.6))

    print(f"文档     : {doc_path}")
    print(f"位置     : {where}" + ("（在表格框里）" if container is not None else "（在框外）"))
    print(f"标签行   : {args.label or '（不加标签）'}")
    print(f"版心     : {cm(box_w)} cm × {cm(box_h)} cm（满宽优先，超高则按高自适应）")
    print(f"排版     : 每张独占一页（图间分页）")
    print("")
    for i, path in enumerate(images, 1):
        w, h = fit_size(path, box_w, box_h)
        print(f"  #{i}  {os.path.basename(path)}  ->  {cm(w)} × {cm(h)} cm"
              + ("   （第一张紧跟标签行）" if i == 1 else "   （前插分页符）"))

    if not apply_mode(args):
        print("\n（未写文件）确认无误后加 --apply 落地；全自动场景用 --auto。")
        return 0

    result = add_attachment(
        doc, images, label=args.label,
        page_break_between=not getattr(args, "no_page_break", False),
        container=container,
    )
    out_path = args.out or _default_out(doc_path)
    trimmed = _trim_trailing_empty_paragraphs(doc)
    doc.save(out_path)
    print(f"\n已写出：{out_path}")
    print(f"       文末追加 {len(result['images'])} 张图片"
          + (f"，标签行「{result['label']}」" if result["label"] else "")
          + (f"；清掉文末 {trimmed} 个空段落（避免多出空白页）" if trimmed else ""))
    return 0


def _default_out(path: str) -> str:
    root, ext = os.path.splitext(path)
    return f"{root}-含附件{ext}"


def apply_mode(args) -> bool:
    return bool(getattr(args, "apply", False) or getattr(args, "auto", False))
