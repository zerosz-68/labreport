"""环境体检：把"缺什么、影响什么"直接打出来。

退出码约定：
    0  核心齐全，可选程序也齐全
    3  核心齐全，但缺可选程序（功能降级，仍可用）
    1  缺核心依赖（必须 pip install）

注：2 留空不用——argparse 在参数用法错误时就返回 2，避免与之混淆。
"""

from __future__ import annotations

import importlib
import json
import platform
import shutil
import sys

# 核心：缺了就跑不了
CORE_MODULES = [
    ("docx", "python-docx", "读写 .docx（必需）"),
    ("lxml", "lxml", "解析 OOXML（必需）"),
    ("PIL", "Pillow", "图片尺寸/OCR 预处理（必需）"),
]

# 可选：缺了只是降级
OPTIONAL_PROGRAMS = [
    ("soffice", "LibreOffice", "转 PDF 目检（没有则跳过排版目检）"),
    ("pandoc", "Pandoc", "LaTeX→OMML 增强（内置转换器可替代）"),
    ("drawio", "draw.io CLI", "流程图导出（没有则交源码）"),
    ("xelatex", "xelatex", "LaTeX 论文编译（没有则只出 docx）"),
    ("pdftoppm", "pdftoppm", "PDF 渲染抽检"),
    ("tesseract", "Tesseract", "本地 OCR 兜底（默认走 agent 视觉，不需要）"),
]


def collect() -> dict:
    py_ok = sys.version_info >= (3, 9)

    core = []
    for module, dist, why in CORE_MODULES:
        try:
            importlib.import_module(module)
            core.append({"module": module, "dist": dist, "why": why, "ok": True})
        except Exception as exc:  # noqa: BLE001
            core.append(
                {
                    "module": module,
                    "dist": dist,
                    "why": why,
                    "ok": False,
                    "error": str(exc),
                }
            )

    optional = []
    for exe, name, why in OPTIONAL_PROGRAMS:
        path = shutil.which(exe)
        optional.append({"name": name, "exe": exe, "path": path, "why": why, "ok": bool(path)})

    system = platform.system()
    return {
        "python": {
            "version": platform.python_version(),
            "executable": sys.executable,
            "ok": py_ok,
            "need": ">=3.9",
        },
        "system": system,
        "core": core,
        "optional": optional,
        "cjk_font_hint": (
            "Windows 自带宋体/黑体/微软雅黑，通常无需处理"
            if system == "Windows"
            else "macOS/Linux 建议装 Noto Sans CJK / Source Han Sans，否则图表与 PDF 中文可能缺字"
        ),
    }


def run(args) -> int:
    data = collect()

    if getattr(args, "json", False):
        print(json.dumps(data, ensure_ascii=False, indent=2))
    else:
        print(f"labreport doctor —— 环境体检")
        print(f"  系统      : {data['system']}")
        print(f"  Python    : {data['python']['version']}（要求 {data['python']['need']}）")
        print(f"  解释器路径 : {data['python']['executable']}")
        print()
        print("  [核心依赖]")
        for item in data["core"]:
            mark = "OK  " if item["ok"] else "缺失"
            extra = "" if item["ok"] else f"  ->  pip install {item['dist']}"
            print(f"    {mark} {item['dist']:<14} {item['why']}{extra}")
        print()
        print("  [可选程序：缺失只降级，不阻塞]")
        for item in data["optional"]:
            mark = "OK  " if item["ok"] else "缺  "
            where = item["path"] or item["why"]
            print(f"    {mark} {item['name']:<16} {where}")
        print()
        print(f"  字体提示  : {data['cjk_font_hint']}")

    core_missing = [c for c in data["core"] if not c["ok"]] or not data["python"]["ok"]
    opt_missing = [o for o in data["optional"] if not o["ok"]]
    if core_missing:
        return 1
    return 3 if opt_missing else 0
