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
from docx.shared import Emu

EMU_PER_CM = 360000


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


def add_attachment(doc, images, label: str = "附件", page_break_between: bool = True,
                   align_center: bool = True) -> dict:
    """在文档末尾追加「标签行 + 每页一张的图片」。"""
    box_w, box_h = usable_box(doc)
    added = []

    if label:
        p = doc.add_paragraph()
        p.add_run(label)

    first = True
    for path in images:
        w, h = fit_size(path, box_w, box_h)
        p = doc.add_paragraph()
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

    print(f"文档     : {doc_path}")
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
    )
    out_path = args.out or _default_out(doc_path)
    doc.save(out_path)
    print(f"\n已写出：{out_path}")
    print(f"       文末追加 {len(result['images'])} 张图片"
          + (f"，标签行「{result['label']}」" if result["label"] else ""))
    return 0


def _default_out(path: str) -> str:
    root, ext = os.path.splitext(path)
    return f"{root}-含附件{ext}"


def apply_mode(args) -> bool:
    return bool(getattr(args, "apply", False) or getattr(args, "auto", False))
