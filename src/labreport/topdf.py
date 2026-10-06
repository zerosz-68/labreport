"""docx → PDF（目检用），可选把 PDF 渲染成 PNG 让 agent 亲眼看排版。

转换器按可用性依次尝试：
1. **Microsoft Word（COM，仅 Windows）**——本机已装 Office 时最保真；
2. **LibreOffice（soffice）**——跨平台；
3. 都没有：返回"降级"（退出码 3），并告诉用户装什么，不影响其余命令。

子进程一律把 stdout/stderr 重定向到**文件**（不用管道），避免受限环境下的管道限制。
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys

WORD_SCRIPT = r"""$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try { $word.DisplayAlerts = 0 } catch { }
try {
  $doc = $word.Documents.Open('__DOCX__')
  try {
    $doc.ExportAsFixedFormat('__PDF__', 17)
  } catch {
    $doc.SaveAs([ref]'__PDF__', [ref]17)
  }
  $doc.Close(0)
} finally {
  # Word 有时在导出后自身异常退出（受限环境下常见），此时 PDF 已经写好；
  # 不让 Quit 的异常把整个脚本判成失败。
  try { $word.Quit() } catch { }
}
"""


def _log_path(outdir: str, name: str) -> str:
    os.makedirs(outdir, exist_ok=True)
    return os.path.join(outdir, name)


def _run(cmd, logfile, timeout=300):
    with open(logfile, "w", encoding="utf-8") as log:
        try:
            proc = subprocess.run(
                cmd, stdout=log, stderr=subprocess.STDOUT, timeout=timeout, shell=False
            )
            return proc.returncode
        except Exception as exc:  # noqa: BLE001
            log.write(f"\n[launcher] {exc}\n")
            return -1


def find_soffice() -> str | None:
    found = shutil.which("soffice") or shutil.which("soffice.exe")
    if found:
        return found
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/bin/soffice",
        "/usr/local/bin/soffice",
    ]
    for c in candidates:
        if os.path.isfile(c):
            return c
    return None


def word_available() -> bool:
    if sys.platform != "win32":
        return False
    return shutil.which("powershell") is not None or shutil.which("pwsh") is not None


def convert_with_word(docx: str, pdf: str, workdir: str) -> tuple[bool, str]:
    shell = shutil.which("powershell") or shutil.which("pwsh")
    script = WORD_SCRIPT.replace("__DOCX__", docx.replace("'", "''")).replace(
        "__PDF__", pdf.replace("'", "''")
    )
    spath = _log_path(workdir, "word_convert.ps1")
    with open(spath, "w", encoding="utf-8-sig") as fh:
        fh.write(script)
    log = _log_path(workdir, "word_convert.log")
    code = _run([shell, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", spath], log)
    ok = os.path.isfile(pdf) and os.path.getsize(pdf) > 0
    if ok:
        return True, "Word COM"
    msg = f"Word COM（退出码 {code}）"
    if os.path.isfile(log):
        with open(log, encoding="utf-8", errors="ignore") as fh:
            tail = [l for l in fh.read().strip().splitlines() if l.strip()]
            if tail:
                msg += "：" + tail[-1][:200]
    return False, msg


def convert_with_soffice(docx: str, outdir: str, workdir: str) -> tuple[bool, str]:
    soffice = find_soffice()
    if not soffice:
        return False, "未找到 LibreOffice"
    log = _log_path(workdir, "soffice_convert.log")
    code = _run(
        [soffice, "--headless", "--norestore", "--convert-to", "pdf", "--outdir", outdir, docx],
        log,
    )
    expect = os.path.join(outdir, os.path.splitext(os.path.basename(docx))[0] + ".pdf")
    ok = os.path.isfile(expect) and os.path.getsize(expect) > 0
    return ok, f"LibreOffice（退出码 {code}）"


def render_pages(pdf: str, outdir: str, dpi: int = 110, first: int = 1, last: int | None = None,
                 workdir: str | None = None) -> tuple[bool, list[str], str]:
    exe = shutil.which("pdftoppm")
    if not exe:
        return False, [], "未找到 pdftoppm（Poppler；MiKTeX/TeX Live 通常自带）"
    os.makedirs(outdir, exist_ok=True)
    prefix = os.path.join(outdir, "page")
    cmd = [exe, "-png", "-r", str(dpi), "-f", str(first)]
    if last:
        cmd += ["-l", str(last)]
    cmd += [pdf, prefix]
    log = _log_path(workdir or outdir, "pdftoppm.log")
    code = _run(cmd, log)
    pngs = sorted(
        os.path.join(outdir, f) for f in os.listdir(outdir) if f.startswith("page") and f.endswith(".png")
    )
    return (code == 0 and bool(pngs)), pngs, f"pdftoppm（退出码 {code}）"


def run(args) -> int:
    target = args.target
    if not os.path.isfile(target):
        print(f"找不到文件：{target}")
        return 1
    if not target.lower().endswith(".docx"):
        print("只支持 .docx 输入")
        return 1

    out_pdf = args.out or os.path.splitext(target)[0] + ".pdf"
    workdir = os.path.join(os.path.dirname(os.path.abspath(out_pdf)), "_labreport_pdf_logs")

    print(f"输入 : {target}")
    print(f"输出 : {out_pdf}")

    tried = []
    ok = False
    if word_available():
        ok, msg = convert_with_word(os.path.abspath(target), os.path.abspath(out_pdf), workdir)
        tried.append(msg)
        if ok:
            print(f"转换 : ✓ {msg}")
    if not ok:
        ok, msg = convert_with_soffice(os.path.abspath(target), os.path.dirname(os.path.abspath(out_pdf)), workdir)
        tried.append(msg)
        if ok:
            print(f"转换 : ✓ {msg}")
            produced = os.path.join(
                os.path.dirname(os.path.abspath(out_pdf)),
                os.path.splitext(os.path.basename(target))[0] + ".pdf",
            )
            if os.path.abspath(produced) != os.path.abspath(out_pdf) and os.path.isfile(produced):
                shutil.move(produced, out_pdf)

    if not ok:
        print("转换 : ✗ 没有可用的转换器")
        for t in tried:
            print(f"       - {t}")
        print("\n装一个即可（不影响其它命令）：")
        print("  Windows : 安装 Microsoft Word，或 LibreOffice")
        print("  macOS   : brew install --cask libreoffice")
        print("  Linux   : sudo apt install libreoffice    # 或 dnf install libreoffice")
        return 3

    size = round(os.path.getsize(out_pdf) / 1024, 1)
    print(f"       PDF {size} KB")

    if getattr(args, "pages", False):
        outdir = args.outdir or os.path.splitext(out_pdf)[0] + "_pages"
        ok2, pngs, msg = render_pages(
            out_pdf, outdir, int(args.dpi), int(getattr(args, "first", 1) or 1),
            getattr(args, "last", None), workdir,
        )
        print(f"渲染 : {'✓' if ok2 else '✗'} {msg}")
        if ok2:
            print(f"       共 {len(pngs)} 张 PNG -> {outdir}")
            for p in pngs[:6]:
                print(f"       {p}")
            print("       用看图工具/视觉能力逐页检查：中文字体、公式显示、表格不破版。")
    return 0
