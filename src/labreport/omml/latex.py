"""内置 LaTeX → OMML（Word 原生公式）转换，不依赖 pandoc。

支持范围（覆盖实验报告里实际会出现的写法）：
    分数 \\frac{}{}        根号 \\sqrt{} / \\sqrt[n]{}     上下标 ^ _ （可同时）
    括号 \\left( \\right)  正体 \\mathrm{} \\text{}         求和/积分 \\sum \\int（带上下限）
    重音 \\bar{} \\vec{}   希腊字母、常见运算符、\\, \\; 空白
输出风格对齐 pandoc 生成的 OMML：变量为普通 run，运算符/单位为正体 run
(`<m:rPr><m:sty m:val="p"/></m:rPr>`)，分数带 `<m:fPr><m:type m:val="bar"/></m:fPr>`。
"""

from __future__ import annotations

import json
import os
import re
import xml.etree.ElementTree as ET

M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

NSDECL = (
    f'xmlns:m="{M}" '
    f'xmlns:w="{W}" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:o="urn:schemas-microsoft-com:office:office" '
    'xmlns:v="urn:schemas-microsoft-com:vml" '
    'xmlns:w10="urn:schemas-microsoft-com:office:word" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
)

# 小写希腊字母：Word 里按变量处理（斜体），与 pandoc 一致
GREEK = {
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", "epsilon": "ε",
    "varepsilon": "ε", "zeta": "ζ", "eta": "η", "theta": "θ", "iota": "ι",
    "kappa": "κ", "lambda": "λ", "mu": "μ", "nu": "ν", "xi": "ξ", "pi": "π",
    "rho": "ρ", "sigma": "σ", "tau": "τ", "upsilon": "υ", "phi": "φ",
    "varphi": "φ", "chi": "χ", "psi": "ψ", "omega": "ω",
}
GREEK_UPPER = {
    "Gamma": "Γ", "Delta": "Δ", "Theta": "Θ", "Lambda": "Λ", "Xi": "Ξ",
    "Pi": "Π", "Sigma": "Σ", "Upsilon": "Υ", "Phi": "Φ", "Psi": "Ψ", "Omega": "Ω",
}
# 运算符/单位：正体（m:sty=p）
OPERATORS = {
    "times": "×", "cdot": "⋅", "div": "÷", "pm": "±", "mp": "∓",
    "approx": "≈", "propto": "∝", "le": "≤", "leq": "≤", "ge": "≥", "geq": "≥",
    "ne": "≠", "neq": "≠", "equiv": "≡", "sim": "∼", "infty": "∞",
    "partial": "∂", "circ": "∘", "degree": "°", "angle": "∠", "perp": "⊥",
    "parallel": "∥", "therefore": "∴", "because": "∵", "rightarrow": "→",
    "to": "→", "leftarrow": "←", "leftrightarrow": "↔", "ldots": "…", "dots": "…",
    "quad": "\u2003", "qquad": "\u2003\u2003",
}
SPACES = {",": "\u2009", ";": "\u2005", ":": "\u2005", "!": "", " ": " "}
DELIMS = {"(": "(", ")": ")", "[": "[", "]": "]", "|": "|", ".": "", "{": "{", "}": "}"}
NARY_CHR = {"sum": "∑", "prod": "∏", "int": "∫", "oint": "∮"}


def _esc(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


# ---------------------------------------------------------------- tokenizer
def tokenize(src: str):
    toks, i, n = [], 0, len(src)
    while i < n:
        c = src[i]
        if c.isspace():
            j = i
            while j < n and src[j].isspace():
                j += 1
            toks.append(("space", " "))
            i = j
            continue
        if c == "\\":
            j = i + 1
            if j < n and src[j].isalpha():
                while j < n and src[j].isalpha():
                    j += 1
                toks.append(("cmd", src[i + 1 : j]))
            elif j < n:
                toks.append(("sym", src[j]))
                j += 1
            else:
                j = n
            i = j
            continue
        if c in "{}^_[]":
            toks.append((c, c))
            i += 1
            continue
        if c.isdigit():
            j = i
            while j < n and (src[j].isdigit() or src[j] == "."):
                j += 1
            toks.append(("num", src[i:j]))
            i = j
            continue
        toks.append(("char", c))
        i += 1
    return toks


# ---------------------------------------------------------------- parser
class Parser:
    """把 token 流解析成中间节点树。节点为元组，第一个元素是类型。"""

    def __init__(self, toks):
        self.toks = toks
        self.i = 0

    def peek(self):
        return self.toks[self.i] if self.i < len(self.toks) else (None, None)

    def advance(self):
        tok = self.peek()
        self.i += 1
        return tok

    def skip_spaces(self):
        while self.peek()[0] == "space":
            self.i += 1

    def parse_seq(self, stops=()):
        out = []
        while True:
            kind, val = self.peek()
            if kind is None or kind in stops:
                break
            if kind == "char" and val in stops:
                break
            if kind == "cmd" and val == "right":
                break
            if kind == "space":
                self.i += 1
                continue
            node = self.parse_postfix()
            if node is None:
                break
            out.append(node)
        return out

    def parse_arg(self):
        self.skip_spaces()
        kind, _ = self.peek()
        if kind == "{":
            self.advance()
            nodes = self.parse_seq(stops=("}",))
            if self.peek()[0] == "}":
                self.advance()
            return nodes
        node = self.parse_atom()
        return [node] if node is not None else []

    def parse_postfix(self):
        base = self.parse_atom()
        if base is None:
            return None
        sub = sup = None
        while True:
            self.skip_spaces()
            kind, _ = self.peek()
            if kind in ("^", "_"):
                self.advance()
                arg = self.parse_arg()
                if kind == "^":
                    sup = arg
                else:
                    sub = arg
            else:
                break
        if sub is not None and sup is not None:
            return ("sSubSup", base, sub, sup)
        if sub is not None:
            return ("sSub", base, sub)
        if sup is not None:
            return ("sSup", base, sup)
        return base

    def parse_atom(self):
        self.skip_spaces()
        kind, val = self.advance()
        if kind is None:
            return None
        if kind == "{":
            nodes = self.parse_seq(stops=("}",))
            if self.peek()[0] == "}":
                self.advance()
            return ("seq", nodes)
        if kind == "char" and val == "(":
            # 普通圆括号也生成 m:d 定界符（与 pandoc 一致），随后 ^/_ 会挂在定界符上
            body = self.parse_seq(stops=(")",))
            if self.peek() == ("char", ")"):
                self.advance()
            return ("delim", "(", ")", body)
        if kind == "[":
            body = self.parse_seq(stops=("]",))
            if self.peek()[0] == "]":
                self.advance()
            return ("delim", "[", "]", body)
        if kind == "char":
            if val in "=+-*/<>≤≥≠±,;:!%|'":
                mapped = {"-": "−", "*": "∗"}.get(val, val)
                return ("run", mapped, True)
            if val == ")":
                return ("run", ")", True)
            return ("run", val, False)
        if kind == "num":
            return ("run", val, False)
        if kind == "sym":
            if val in SPACES:
                # \, \; \! 等空白：pandoc 生成的是 U+2009 且不带正体标记
                return ("run", SPACES[val], False)
            if val in "{}%$#&_":
                return ("run", val, True)
            return ("run", val, True)
        if kind == "cmd":
            return self.parse_command(val)
        return None

    # ---- 具体命令
    def parse_command(self, name):
        if name == "frac" or name == "dfrac" or name == "tfrac":
            num = self.parse_arg()
            den = self.parse_arg()
            return ("frac", num, den)
        if name == "sqrt":
            self.skip_spaces()
            deg = None
            if self.peek()[0] == "[":
                self.advance()
                deg = self.parse_seq(stops=("]",))
                if self.peek()[0] == "]":
                    self.advance()
            body = self.parse_arg()
            return ("rad", body, deg)
        if name in ("text", "mbox", "textrm"):
            # \text 按原样文本输出（保留空格），正体单 run
            return ("run", self.raw_string(), True)
        if name in ("mathrm", "mathbf", "mathit", "operatorname", "textnormal"):
            # \mathrm 内部仍按数学解析（上标/分数等结构保留），但所有 run 强制正体
            body = self.parse_arg()
            return ("seq", [_force_plain(n) for n in body])
        if name == "left":
            self.skip_spaces()
            k, v = self.advance()
            beg = DELIMS.get(v, v)
            body = self.parse_seq()
            end = ""
            if self.peek()[0] == "cmd" and self.peek()[1] == "right":
                self.advance()
                self.skip_spaces()
                k2, v2 = self.advance()
                end = DELIMS.get(v2, v2)
            return ("delim", beg, end, body)
        if name in NARY_CHR:
            sub = sup = None
            while True:
                self.skip_spaces()
                kind, _ = self.peek()
                if kind in ("_", "^"):
                    self.advance()
                    arg = self.parse_arg()
                    if kind == "_":
                        sub = arg
                    else:
                        sup = arg
                else:
                    break
            self.skip_spaces()
            body = self.parse_arg()
            return ("nary", NARY_CHR[name], sub, sup, body)
        if name in ("bar", "overline"):
            return ("acc", "\u0304", self.parse_arg())
        if name == "vec":
            return ("acc", "\u20d7", self.parse_arg())
        if name == "hat":
            return ("acc", "\u0302", self.parse_arg())
        if name == "dot":
            return ("acc", "\u0307", self.parse_arg())
        if name in OPERATORS:
            return ("run", OPERATORS[name], True)
        if name in GREEK:
            return ("run", GREEK[name], False)
        if name in GREEK_UPPER:
            return ("run", GREEK_UPPER[name], True)
        if name in ("left.",):
            return None
        # 未知命令：退化为正体文本，避免整条公式失败
        return ("run", "\\" + name, True)

    def raw_string(self):
        """读取 {..} 内的原始文本（保留空格），供 \\text 使用。"""
        self.skip_spaces()
        if self.peek()[0] != "{":
            return ""
        self.advance()
        depth, parts = 1, []
        while self.i < len(self.toks):
            kind, val = self.advance()
            if kind == "{":
                depth += 1
                parts.append("{")
                continue
            if kind == "}":
                depth -= 1
                if depth == 0:
                    break
                parts.append("}")
                continue
            if kind == "space":
                parts.append(" ")
                continue
            if kind == "cmd":
                ch = (
                    OPERATORS.get(val)
                    or GREEK.get(val)
                    or GREEK_UPPER.get(val)
                    or SPACES.get(val)
                )
                parts.append(ch if ch is not None else "\\" + val)
                continue
            if kind == "sym":
                parts.append(SPACES.get(val) or val)
                continue
            parts.append(val)
        return "".join(parts)


def _force_plain(node):
    """把子树里所有 run 标记为正体（用于 \\mathrm{...}），结构节点保持不变。"""
    if node is None:
        return None
    kind = node[0]
    rest = []
    for item in node[1:]:
        if isinstance(item, list):
            rest.append([_force_plain(x) for x in item])
        elif isinstance(item, tuple):
            rest.append(_force_plain(item))
        else:
            rest.append(item)
    if kind == "run":
        return ("run", rest[0], True)
    return tuple([kind] + rest)


# ---------------------------------------------------------------- renderer
def _render(node) -> str:
    if node is None:
        return ""
    kind = node[0]
    if kind == "seq":
        return "".join(_render(n) for n in node[1])
    if kind == "run":
        text, plain = node[1], node[2]
        if text == "":
            return ""
        pr = '<m:rPr><m:sty m:val="p"/></m:rPr>' if plain else ""
        return f"<m:r>{pr}<m:t>{_esc(text)}</m:t></m:r>"
    if kind == "text":
        return "".join(_render(n) for n in node[1])
    if kind == "frac":
        return (
            "<m:f><m:fPr><m:type m:val=\"bar\"/></m:fPr>"
            f"<m:num>{_render_seq(node[1])}</m:num>"
            f"<m:den>{_render_seq(node[2])}</m:den></m:f>"
        )
    if kind == "rad":
        body, deg = node[1], node[2]
        if deg:
            return (
                f"<m:rad><m:deg>{_render_seq(deg)}</m:deg>"
                f"<m:e>{_render_seq(body)}</m:e></m:rad>"
            )
        return (
            '<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/>'
            f"<m:e>{_render_seq(body)}</m:e></m:rad>"
        )
    if kind == "sSup":
        return f"<m:sSup><m:e>{_render(node[1])}</m:e><m:sup>{_render_seq(node[2])}</m:sup></m:sSup>"
    if kind == "sSub":
        return f"<m:sSub><m:e>{_render(node[1])}</m:e><m:sub>{_render_seq(node[2])}</m:sub></m:sSub>"
    if kind == "sSubSup":
        return (
            f"<m:sSubSup><m:e>{_render(node[1])}</m:e>"
            f"<m:sub>{_render_seq(node[2])}</m:sub>"
            f"<m:sup>{_render_seq(node[3])}</m:sup></m:sSubSup>"
        )
    if kind == "delim":
        beg, end, body = node[1], node[2], node[3]
        return (
            "<m:d><m:dPr>"
            f'<m:begChr m:val="{_esc(beg)}"/><m:endChr m:val="{_esc(end)}"/>'
            '<m:sepChr m:val=""/><m:grow/></m:dPr>'
            f"<m:e>{_render_seq(body)}</m:e></m:d>"
        )
    if kind == "nary":
        chr_, sub, sup, body = node[1], node[2], node[3], node[4]
        pr = f'<m:naryPr><m:chr m:val="{_esc(chr_)}"/><m:limLoc m:val="undOvr"/>'
        pr += '<m:subHide m:val="0"/>' if sub else '<m:subHide m:val="1"/>'
        pr += '<m:supHide m:val="0"/>' if sup else '<m:supHide m:val="1"/>'
        pr += "</m:naryPr>"
        return (
            f"<m:nary>{pr}"
            f"<m:sub>{_render_seq(sub or [])}</m:sub>"
            f"<m:sup>{_render_seq(sup or [])}</m:sup>"
            f"<m:e>{_render_seq(body)}</m:e></m:nary>"
        )
    if kind == "acc":
        return (
            f'<m:acc><m:accPr><m:chr m:val="{_esc(node[1])}"/></m:accPr>'
            f"<m:e>{_render_seq(node[2])}</m:e></m:acc>"
        )
    return ""


def _render_seq(nodes) -> str:
    return "".join(_render(n) for n in (nodes or []))


FRAG_RE = re.compile(r"^\s*\$?(.*?)\$?\s*$")


def latex_to_omml(latex: str) -> str:
    """把一条 LaTeX 公式转成完整的 <m:oMath> 片段字符串。"""
    src = latex.strip()
    if src.startswith("$") and src.endswith("$"):
        src = src[1:-1]
    nodes = Parser(tokenize(src)).parse_seq()
    return f"<m:oMath {NSDECL}>{_render_seq(nodes)}</m:oMath>"


# ---------------------------------------------------------------- 结构比对
def _sig(el):
    """把一个 XML 元素压成可比较的签名（忽略命名空间前缀与空白）。"""
    tag = el.tag.split("}")[-1]
    attrs = tuple(sorted((k.split("}")[-1], v) for k, v in el.attrib.items()))
    text = (el.text or "").strip()
    kids = tuple(_sig(c) for c in el)
    return (tag, attrs, text, kids)


def compare(expected_xml: str, actual_xml: str) -> dict:
    exp = ET.fromstring(expected_xml)
    act = ET.fromstring(actual_xml)
    e, a = _sig(exp), _sig(act)
    return {"equal": e == a, "expected": e, "actual": a}


def _first_diff(e, a, path="oMath"):
    if e[0] != a[0]:
        return f"{path}: 标签不同 expected=<{e[0]}> actual=<{a[0]}>"
    if e[1] != a[1]:
        return f"{path}: 属性不同 expected={e[1]} actual={a[1]}"
    if e[2] != a[2]:
        return f"{path}: 文本不同 expected={e[2]!r} actual={a[2]!r}"
    if len(e[3]) != len(a[3]):
        return f"{path}: 子节点数不同 expected={len(e[3])} actual={len(a[3])}"
    for i, (ec, ac) in enumerate(zip(e[3], a[3])):
        d = _first_diff(ec, ac, f"{path}/{ec[0]}[{i}]")
        if d:
            return d
    return ""


# ---------------------------------------------------------------- CLI
def run(args) -> int:
    path = args.formulas
    if not os.path.isfile(path):
        print(f"找不到公式文件：{path}")
        return 1

    lines = []
    with open(path, encoding="utf-8") as fh:
        for lineno, raw in enumerate(fh, 1):
            s = raw.strip()
            if not s or s.startswith("#"):
                continue
            lines.append((lineno, s))

    outdir = args.out or os.path.join(os.path.dirname(os.path.abspath(path)), "out")
    os.makedirs(outdir, exist_ok=True)

    results, fails = [], []
    for idx, (lineno, latex) in enumerate(lines, 1):
        try:
            xml = latex_to_omml(latex)
            ET.fromstring(xml)  # 自校验：必须是合法 XML
        except Exception as exc:  # noqa: BLE001
            fails.append((idx, lineno, latex, str(exc)))
            continue
        name = f"f{idx:02d}.xml"
        with open(os.path.join(outdir, name), "w", encoding="utf-8") as fh:
            fh.write(xml)
        results.append({"n": idx, "line": lineno, "latex": latex, "xml": xml, "file": name})

    with open(os.path.join(outdir, "frags.json"), "w", encoding="utf-8") as fh:
        json.dump(results, fh, ensure_ascii=False, indent=1)

    if args.stdout:
        want = int(args.stdout)
        for r in results:
            if r["n"] == want:
                print(r["xml"])
                return 0
        print(f"没有第 {want} 条公式")
        return 1

    print(f"转换完成：{len(results)} 条 -> {outdir}")
    if fails:
        print(f"失败 {len(fails)} 条：")
        for idx, lineno, latex, err in fails[:10]:
            print(f"  #{idx} (第{lineno}行) {latex[:50]} -> {err}")

    if args.compare:
        cmp_dir = args.compare
        ok = 0
        details = []
        for r in results:
            exp_file = os.path.join(cmp_dir, r["file"])
            if not os.path.isfile(exp_file):
                details.append((r["n"], "缺对照", ""))
                continue
            with open(exp_file, encoding="utf-8") as fh:
                exp_xml = fh.read()
            try:
                res = compare(exp_xml, r["xml"])
            except Exception as exc:  # noqa: BLE001
                details.append((r["n"], "解析失败", str(exc)))
                continue
            if res["equal"]:
                ok += 1
                details.append((r["n"], "一致", ""))
            else:
                details.append((r["n"], "不同", _first_diff(res["expected"], res["actual"])))
        total = len(results)
        print(f"\n与 pandoc 对照：{ok}/{total} 条结构完全一致")
        for n, status, why in details:
            if status != "一致":
                print(f"  #{n:<3} {status}  {why[:150]}")

    return 0
