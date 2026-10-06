#!/usr/bin/env python
"""labreport 回归测试：把每条命令跑一遍并断言退出码。

用法：
    python tests/run_all.py
    python tests/run_all.py --docx-dir "C:/path/to/扭摆实验资料"   # 额外跑需要 .docx 的用例
    python tests/run_all.py --frags-dir "…/脚本与中间文件/omml/frags"  # 额外与 pandoc 参照比对

说明：
- 需要 .docx 的用例（inspect/fill/audit/pdf）在找不到文件时**自动跳过**并说明原因，
  这样仓库里不必放真实报告，别人也能跑通不依赖 docx 的部分。
- 全部在进程内调用 CLI（不经子进程/管道），因此在受限环境下也能跑。
"""

from __future__ import annotations

import argparse
import contextlib
import io
import os
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
FIX = os.path.join(HERE, "fixtures")
OUT = os.path.join(HERE, "_out")
sys.path.insert(0, os.path.join(ROOT, "src"))

from labreport.cli import main as cli_main  # noqa: E402

PASS, FAIL, SKIP = [], [], []


def call(argv):
    """在进程内跑一条命令，返回 (退出码, 输出)。"""
    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf):
            code = cli_main(argv)
    except SystemExit as exc:  # argparse 用法错误
        code = int(exc.code or 0)
    return code, buf.getvalue()


def case(name, argv, expect, contains=None, skip_reason=None):
    if skip_reason:
        SKIP.append((name, skip_reason))
        print(f"  [跳过] {name:<38} {skip_reason}")
        return None
    code, out = call(argv)
    ok = code in expect if isinstance(expect, (tuple, list, set)) else code == expect
    if ok and contains:
        missing = [c for c in contains if c not in out]
        if missing:
            ok = False
            out += f"\n(缺少预期文本: {missing})"
    if ok:
        PASS.append(name)
        print(f"  [OK] {name:<38} 退出码 {code}")
    else:
        FAIL.append((name, argv, code, expect, out[-400:]))
        print(f"  [!!] {name:<38} 退出码 {code}（期望 {expect}）")
    return out


def _make_output_safe() -> None:
    """同 CLI：英文版 Windows（cp1252）下 print 中文不崩。

    测试脚本自己也要打印中文小标题，而且它在调用 CLI 之前就会打印。
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


def main():
    _make_output_safe()
    ap = argparse.ArgumentParser()
    ap.add_argument("--docx-dir", default=os.environ.get("LABREPORT_FIXTURES", ""),
                    help="含 原始资料/2 扭摆实验报告.docx 与 成品/…已填写.docx 的目录")
    ap.add_argument("--frags-dir", default="", help="pandoc 生成的 OMML 参照片段目录")
    args = ap.parse_args()

    shutil.rmtree(OUT, ignore_errors=True)
    os.makedirs(OUT, exist_ok=True)

    tpl = filled = None
    if args.docx_dir:
        tpl = os.path.join(args.docx_dir, "原始资料", "2 扭摆实验报告.docx")
        filled = os.path.join(args.docx_dir, "成品", "2 扭摆实验报告-已填写.docx")
        if not os.path.isfile(tpl):
            tpl = None
        if not os.path.isfile(filled):
            filled = None
    no_docx = "未提供 --docx-dir"

    print("== 1. 环境与公式 ==")
    case("doctor（0 就绪 / 3 缺可选）", ["doctor"], (0, 3))
    case("omml 转换 28 条真实公式", ["omml", os.path.join(FIX, "formulas.md"), "-o", os.path.join(OUT, "omml")], 0,
         contains=["28 条"])
    if args.frags_dir and os.path.isdir(args.frags_dir):
        case("omml 与 pandoc 参照逐条比对", ["omml", os.path.join(FIX, "formulas.md"), "-o",
                                            os.path.join(OUT, "omml2"), "--compare", args.frags_dir], 0,
             contains=["28/28"])
    else:
        SKIP.append(("omml 与 pandoc 参照比对", "未提供 --frags-dir"))
        print(f"  [跳过] {'omml 与 pandoc 参照比对':<38} 未提供 --frags-dir")
    case("omml 语料外公式不退化", ["omml", os.path.join(FIX, "stress.md"), "-o", os.path.join(OUT, "stress")], 0,
         contains=["12 条"])

    print("== 2. 数据校验 ==")
    case("data template", ["data", "template"], 0)
    case("data check 正确数据", ["data", "check", os.path.join(FIX, "torsion.json")], 0,
         contains=["可以填报告"])
    case("data check 抄错版（应报错）", ["data", "check", os.path.join(FIX, "torsion_bad.json")], 1,
         contains=["error"])

    print("== 3. 结构 / 填写 / 自检 ==")
    case("inspect 空白模板", ["inspect", tpl], 0, contains=["表格"], skip_reason=None if tpl else no_docx)
    case("inspect 已填写成品", ["inspect", filled], 0, contains=["公式"], skip_reason=None if filled else no_docx)
    out_filled = os.path.join(OUT, "filled.docx")
    case("fill 默认只出计划", ["fill", tpl, "--map", os.path.join(FIX, "torsion_map.json"),
                             "--data", os.path.join(FIX, "torsion.json"), "-o", out_filled], 0,
         contains=["未写文件"], skip_reason=None if tpl else no_docx)
    if tpl and not os.path.exists(out_filled):
        PASS.append("fill 默认模式未写文件")
        print(f"  [OK] {'fill 默认模式未写文件':<38} 已核实")
    case("fill --auto 落地", ["fill", tpl, "--map", os.path.join(FIX, "torsion_map.json"),
                            "--data", os.path.join(FIX, "torsion.json"), "-o", out_filled, "--auto"], 0,
         contains=["已写出"], skip_reason=None if tpl else no_docx)
    case("audit 自检成品（真实报告应可交付）", ["audit", filled, "--source", tpl, "--data",
                                            os.path.join(FIX, "torsion.json")], 0,
         contains=["可以交付"], skip_reason=None if (tpl and filled) else no_docx)
    case("audit 自检空白模板（应报大量未填）", ["audit", tpl, "--data", os.path.join(FIX, "torsion.json")], 0,
         contains=["数据回读"], skip_reason=None if tpl else no_docx)
    # 注意：out_filled 是路径字符串（恒为真），必须判断"文件是否真的存在"，
    # 否则跳过 fill 之后会拿着不存在的文件去跑 pdf，导致失败。
    pdf_target = out_filled if os.path.isfile(out_filled) else tpl
    case("pdf 转 PDF（无转换器时降级 3）", ["pdf", pdf_target, "-o", os.path.join(OUT, "out.pdf")],
         (0, 3), skip_reason=None if pdf_target else no_docx)

    print("== 4. 适配页与知识卡 ==")
    case("install-skill --list", ["install-skill", "--list"], 0, contains=["claude", "codex"])
    for agent in ("claude", "dsh", "codex", "cursor"):
        case(f"install-skill --print {agent}", ["install-skill", "--for", agent, "--print"], 0,
             contains=["labreport"])
    case("card list（内置卡）", ["card", "list"], 0, contains=["torsion-pendulum"])
    case("card show 内置卡", ["card", "show", "torsion-pendulum"], 0, contains=["平行轴"])
    carddir = os.path.join(OUT, "cards")
    os.environ["LABREPORT_CARD_DIR"] = carddir
    case("card new（自建卡）", ["card", "new", "回归测试卡"], 0, contains=["已从模板创建"])
    case("card new 重复应拒绝", ["card", "new", "回归测试卡"], 1, contains=["已存在"])

    print("\n================ 汇总 ================")
    print(f"通过 {len(PASS)} | 失败 {len(FAIL)} | 跳过 {len(SKIP)}")
    for name, argv, code, expect, tail in FAIL:
        print(f"\n--- 失败：{name}\n    命令：labreport {' '.join(map(str, argv))}"
              f"\n    退出码 {code}，期望 {expect}\n    输出尾部：{tail}")
    for name, why in SKIP:
        print(f"  （跳过 {name}：{why}）")
    return 1 if FAIL else 0


if __name__ == "__main__":
    raise SystemExit(main())
