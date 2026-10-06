"""labreport 命令行入口。

用法：
    labreport doctor                 环境体检（依赖 / 外部程序 / 字体）
    labreport inspect <模板.docx>     读模板结构（表格物理网格、合并、公式、图片）
    labreport --version
"""

from __future__ import annotations

import argparse
import sys

from . import __version__


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="labreport",
        description="通用实验报告 CLI：读模板 / 校验数据 / 填公式 / 自检 / 转 PDF",
    )
    parser.add_argument(
        "-V", "--version", action="version", version=f"labreport {__version__}"
    )
    sub = parser.add_subparsers(dest="cmd", required=True, metavar="<命令>")

    p_doc = sub.add_parser("doctor", help="环境体检：Python 包、外部程序、中文字体")
    p_doc.add_argument("--json", action="store_true", help="以 JSON 输出")

    p_ins = sub.add_parser(
        "inspect", help="读 Word 结构：段落 / 表格物理网格 / 合并单元格 / 公式 / 图片"
    )
    p_ins.add_argument("docx", help="待读取的 .docx 路径")
    p_ins.add_argument("--json", action="store_true", help="以 JSON 输出完整结构")
    p_ins.add_argument(
        "--max-text", type=int, default=18, help="单元格/段落文字截断长度（默认 18）"
    )
    p_ins.add_argument(
        "--formula-sample", type=int, default=8, help="展示前 N 条公式文本（默认 8）"
    )

    p_tbl = sub.add_parser(
        "tables", help="从一份 docx 提取表格（含合并关系）→ 可直接嵌入报告的 JSON"
    )
    p_tbl.add_argument("source", help="源 .docx（如课堂报告书）")
    p_tbl.add_argument("-o", "--out", help="输出表格规格 JSON（不给则只打印摘要）")
    p_tbl.add_argument("--json", action="store_true", help="以 JSON 打印到标准输出")
    p_tbl.add_argument(
        "--keep-title-row", action="store_true", help="保留表格首行的表名（默认把它提为表名）"
    )
    p_tbl.add_argument("--keep-empty", action="store_true", help="保留无数据的空表")

    p_om = sub.add_parser(
        "omml", help="LaTeX → Word 原生公式（OMML）：内置转换，不依赖 pandoc"
    )
    p_om.add_argument("formulas", help="每行一条 LaTeX 的公式清单（可带 $...$）")
    p_om.add_argument("-o", "--out", help="输出目录（默认 <公式文件所在目录>/out）")
    p_om.add_argument(
        "--compare", help="与一个已有 OMML 片段目录逐条做结构比对（如 pandoc 的 frags/）"
    )
    p_om.add_argument("--stdout", help="只打印第 N 条公式的 XML")

    p_data = sub.add_parser(
        "data", help="实验数据校验：agent 读照片后按 data.json 格式交上来，这里验它"
    )
    data_sub = p_data.add_subparsers(dest="data_cmd", metavar="<动作>")
    p_dc = data_sub.add_parser(
        "check", help="校验 data.json：量纲 / 统计量 / 平均值自洽 / 公式重算 / 存疑标注"
    )
    p_dc.add_argument("data_file", help="data.json 路径")
    p_dc.add_argument("--json", action="store_true", help="以 JSON 输出完整报告")
    p_dt = data_sub.add_parser("template", help="输出 data.json 骨架模板（给 agent 照着填）")
    p_dt.add_argument("-o", "--out", help="写入文件（默认打印到屏幕）")

    p_fill = sub.add_parser(
        "fill", help="按物理网格坐标填报告：文本 / 追加 / 清空 / 插入 Word 原生公式"
    )
    p_fill.add_argument("template", help="待填写的 .docx（原模板）")
    p_fill.add_argument("--map", required=True, help="填写映射 map.json")
    p_fill.add_argument("--data", help="data.json（供 {{targets.X}} 等占位符取值）")
    p_fill.add_argument("-o", "--out", help="输出文件（默认 <模板名>-filled.docx）")
    p_fill.add_argument("--apply", action="store_true", help="真正写文件（默认只出计划供确认）")
    p_fill.add_argument("--auto", action="store_true", help="全自动：等同于 --apply")
    p_fill.add_argument("--plan", help="把填写计划另存为 JSON")
    p_fill.add_argument(
        "--tables", help="把 `labreport tables` 提取出的表格 JSON 嵌入正文（配合 --into）"
    )
    p_fill.add_argument("--into", help='嵌入位置 "表,行,列"（1 起，物理网格坐标），如 "2,6,1"')
    p_fill.add_argument(
        "--fit-rows",
        help='把指定行的最小行高改为自适应（只改尺寸，不动边框/字体/合并），如 "2:5,2:6"',
    )

    p_att = sub.add_parser(
        "attach", help="在报告末尾追加附件（标签一行 + 每页一张满宽图片）"
    )
    p_att.add_argument("docx", help="待追加的 .docx（通常是填好的成品）")
    p_att.add_argument("images", nargs="+", help="附件图片，按顺序（如课堂报告书第 1、2 页）")
    p_att.add_argument("--label", default="附件", help='标签行文字（默认"附件"，给空串则不加）')
    p_att.add_argument(
        "--into", help='放进指定单元格（"表,行,列"，1 起）——这样附件在报告框内；不给则追加到文档末尾'
    )
    p_att.add_argument("-o", "--out", help="输出文件（默认 <原名>-含附件.docx）")
    p_att.add_argument("--no-page-break", action="store_true", help="图片之间不插入分页符")
    p_att.add_argument("--apply", action="store_true", help="真正写文件（默认只出计划）")
    p_att.add_argument("--auto", action="store_true", help="全自动：等同于 --apply")

    p_audit = sub.add_parser(
        "audit", help="成品自检：结构 / 公式 / 占位 / 数据回读（交付前最后一道关）"
    )
    p_audit.add_argument("target", help="待检查的成品 .docx")
    p_audit.add_argument("--source", help="对照模板 .docx（检查结构有没有被破坏）")
    p_audit.add_argument("--data", help="data.json（逐项回读数据是否填全）")
    p_audit.add_argument("--tolerance", default=0.002, help="数值命中的相对容差，默认 0.002")
    p_audit.add_argument(
        "--require-attachment", action="store_true", help="要求文末必须有附件（标签 + 图片）"
    )
    p_audit.add_argument("--attachment-label", default="附件", help='附件标签文字（默认"附件"）')
    p_audit.add_argument("--skip-outline", action="store_true", help="跳过“按模板提纲核对章节”")
    p_audit.add_argument("--json", action="store_true", help="以 JSON 输出")

    p_pdf = sub.add_parser(
        "pdf", help="docx → PDF 目检（Word 或 LibreOffice）；--pages 再渲染成 PNG"
    )
    p_pdf.add_argument("target", help="待转换的 .docx")
    p_pdf.add_argument("-o", "--out", help="PDF 路径（默认与 docx 同目录同名）")
    p_pdf.add_argument("--pages", action="store_true", help="把 PDF 每页渲染成 PNG（供视觉目检）")
    p_pdf.add_argument("--outdir", help="PNG 输出目录（默认 <pdf名>_pages/）")
    p_pdf.add_argument("--dpi", type=int, default=110, help="PNG 分辨率，默认 110")
    p_pdf.add_argument("--first", type=int, default=1, help="起始页")
    p_pdf.add_argument("--last", type=int, help="结束页")

    p_sk = sub.add_parser("install-skill", help="把适配页装进 Claude / DSH / Codex / Cursor")
    p_sk.add_argument("--for", dest="for_agent", help="claude | dsh | codex | cursor")
    p_sk.add_argument("--dest", help="自定义落点（默认按 agent 约定）")
    p_sk.add_argument("--print", dest="print_page", action="store_true", help="只打印内容，不写文件")
    p_sk.add_argument("--dry-run", action="store_true", help="只显示将要写入哪里与内容预览")
    p_sk.add_argument("--force", action="store_true", help="目标已存在且不同时覆盖")
    p_sk.add_argument("--list", action="store_true", help="列出支持的 agent 与默认落点")

    p_card = sub.add_parser("card", help="实验知识卡：内置卡 + 自建卡（换实验只换一张卡）")
    card_sub = p_card.add_subparsers(dest="card_cmd", metavar="<动作>")
    card_sub.add_parser("list", help="列出所有知识卡（含来源与摘要）")
    p_cs = card_sub.add_parser("show", help="打印某张知识卡")
    p_cs.add_argument("name", help="卡名（可带或不带 .md）")
    p_cn = card_sub.add_parser("new", help="用模板新建一张知识卡")
    p_cn.add_argument("name", help="卡名，如 用单摆测重力加速度")
    p_cn.add_argument("-o", "--out", help="写入路径（默认 ~/.labreport/cards/<卡名>.md）")
    p_cn.add_argument("--force", action="store_true", help="已存在时覆盖")

    return parser


def _make_output_safe() -> None:
    """让输出在"非 UTF-8 控制台"上也不会崩。

    英文版 Windows 的控制台/管道默认是 cp1252，直接 print 中文会抛
    UnicodeEncodeError 把整个命令打断（GitHub Actions 的 windows runner 就是
    这种情况）。这里分两种：
      - 输出到管道/文件（非终端）：改用 UTF-8（CI 日志、重定向都正常）；
      - 输出到终端：保留系统编码，仅把无法编码的字符降级成 "?"（不崩、不换码）。
    """
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is None:
            continue
        try:
            if getattr(stream, "isatty", lambda: False)():
                reconfigure(errors="replace")
            else:
                reconfigure(encoding="utf-8", errors="replace")
        except Exception:  # noqa: BLE001
            pass


def main(argv=None) -> int:
    _make_output_safe()
    args = build_parser().parse_args(argv)

    if args.cmd == "doctor":
        from .doctor import run as run_doctor

        return run_doctor(args)
    if args.cmd == "inspect":
        from .docx.inspect import run as run_inspect

        return run_inspect(args)
    if args.cmd == "tables":
        from .docx.tables import run as run_tables

        return run_tables(args)
    if args.cmd == "attach":
        from .docx.attach import run as run_attach

        return run_attach(args)
    if args.cmd == "omml":
        from .omml.latex import run as run_omml

        return run_omml(args)
    if args.cmd == "data":
        from .data.check import run as run_data

        return run_data(args)
    if args.cmd == "fill":
        from .docx.fill import run as run_fill

        return run_fill(args)
    if args.cmd == "audit":
        from .audit.check import run as run_audit

        return run_audit(args)
    if args.cmd == "pdf":
        from .topdf import run as run_pdf

        return run_pdf(args)
    if args.cmd == "install-skill":
        from .skills.install import run as run_skill

        return run_skill(args)
    if args.cmd == "card":
        from .card import run as run_card

        return run_card(args)

    print(f"未知命令：{args.cmd}", file=sys.stderr)
    return 2


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
