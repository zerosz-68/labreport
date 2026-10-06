"""数据校验。

职责边界（很重要）：**看照片识别数字是 agent 的活**，本模块只做两件事——
1. 规定一个严格的中间格式（data.json），让 agent 有明确目标可填；
2. 校验这个文件：量纲、统计量、平均值是否自洽、公式重算与预期值比对、存疑标注。

这样切分之后：CLI 不需要视觉能力（任何环境都能跑），而识别质量的锅归 agent 的
视觉能力，校验结果会明确告诉它"哪一格要回看照片"。
"""

from __future__ import annotations

import ast
import json
import math
import operator
import os

# 单位 -> (量纲, 换算到 SI 的系数)
UNITS = {
    "": ("无量纲", 1.0),
    "1": ("无量纲", 1.0),
    "%": ("无量纲", 1.0),
    "mg": ("质量", 1e-6),
    "g": ("质量", 1e-3),
    "kg": ("质量", 1.0),
    "mm": ("长度", 1e-3),
    "cm": ("长度", 1e-2),
    "m": ("长度", 1.0),
    "ms": ("时间", 1e-3),
    "s": ("时间", 1.0),
    "min": ("时间", 60.0),
    "kg*m^2": ("转动惯量", 1.0),
    "g*cm^2": ("转动惯量", 1e-7),
    "N*m": ("力矩", 1.0),
    "N*m/rad": ("扭转系数", 1.0),
    "m/s": ("速度", 1.0),
    "m/s^2": ("加速度", 1.0),
    "m/s2": ("加速度", 1.0),
    "N": ("力", 1.0),
    "J": ("能量", 1.0),
    "Hz": ("频率", 1.0),
    "V": ("电压", 1.0),
    "A": ("电流", 1.0),
    "Ω": ("电阻", 1.0),
    "°C": ("温度", 1.0),
}

_FUNCS = {
    "sqrt": math.sqrt, "abs": abs, "sin": math.sin, "cos": math.cos,
    "tan": math.tan, "log": math.log, "ln": math.log, "exp": math.exp,
    "pi": math.pi, "e": math.e,
}
_BIN = {
    ast.Add: operator.add, ast.Sub: operator.sub, ast.Mult: operator.mul,
    ast.Div: operator.truediv, ast.Pow: operator.pow, ast.Mod: operator.mod,
}
_UNARY = {ast.USub: operator.neg, ast.UAdd: operator.pos}


def safe_eval(expr: str, names: dict) -> float:
    """安全求值：白名单 AST 节点，不用 eval。支持 ^ 作幂（报告里的习惯写法）。"""
    src = expr.strip()
    if "^" in src and "**" not in src:
        src = src.replace("^", "**")
    tree = ast.parse(src, mode="eval")

    def ev(node):
        if isinstance(node, ast.Expression):
            return ev(node.body)
        if isinstance(node, ast.Constant):
            if isinstance(node.value, (int, float)):
                return float(node.value)
            raise ValueError("只允许数字常量")
        if isinstance(node, ast.Name):
            if node.id in names:
                return names[node.id]
            if node.id in _FUNCS:
                return _FUNCS[node.id]
            raise ValueError(f"未知名称：{node.id}")
        if isinstance(node, ast.BinOp):
            return _BIN[type(node.op)](ev(node.left), ev(node.right))
        if isinstance(node, ast.UnaryOp):
            return _UNARY[type(node.op)](ev(node.operand))
        if isinstance(node, ast.Call):
            fn = ev(node.func)
            if not callable(fn):
                raise ValueError("调用了非函数")
            return fn(*[ev(a) for a in node.args])
        raise ValueError(f"不支持的语法：{type(node).__name__}")

    return ev(tree)


def _precision_of(values) -> int:
    """从数值本身推断末位精度（用于平均值自洽判定）。"""
    worst = 0
    for v in values:
        s = repr(float(v))
        if "e" in s or "E" in s:
            continue
        if "." in s:
            worst = max(worst, len(s.split(".")[1].rstrip("0")))
    return worst


def check(data: dict, path: str = "") -> dict:
    issues = []
    names = {"pi": math.pi, "e": math.e}

    # ---------- 常量与量（按 unit 自动归一化到 SI，表达式里直接用量纲一致的数）
    quantities = data.get("quantities", {}) or {}
    q_report = []
    for key, q in quantities.items():
        if not isinstance(q, dict):
            issues.append({"level": "error", "where": f"quantities.{key}", "message": "必须是对象"})
            continue
        if "value" not in q:
            issues.append({"level": "error", "where": f"quantities.{key}", "message": "缺 value"})
            continue
        unit = q.get("unit", "")
        if unit not in UNITS:
            issues.append(
                {"level": "warn", "where": f"quantities.{key}", "message": f"未知单位 {unit!r}（可在 UNITS 表里补充）"}
            )
        factor = UNITS.get(unit, ("", 1.0))[1]
        si_value = float(q["value"]) * factor
        names[key] = si_value
        if q.get("symbol"):
            names[q["symbol"]] = si_value
        if not q.get("source"):
            issues.append({"level": "info", "where": f"quantities.{key}", "message": "没有 source 定位，无法回溯照片"})
        q_report.append(
            {
                "key": key,
                "value": float(q["value"]),
                "unit": unit,
                "valueSI": si_value,
                "factor": factor,
                "source": q.get("source", ""),
            }
        )

    # ---------- 测量组
    stats = []
    measurements = data.get("measurements", []) or []
    seen_ids = set()
    for m in measurements:
        mid = m.get("id")
        if not mid:
            issues.append({"level": "error", "where": "measurements", "message": "有一条缺 id"})
            continue
        if mid in seen_ids:
            issues.append({"level": "error", "where": f"measurements.{mid}", "message": "id 重复"})
        seen_ids.add(mid)
        vals = m.get("values") or []
        if not vals:
            issues.append({"level": "error", "where": f"measurements.{mid}", "message": "values 为空"})
            continue
        try:
            vals = [float(v) for v in vals]
        except (TypeError, ValueError):
            issues.append({"level": "error", "where": f"measurements.{mid}", "message": "values 含非数字"})
            continue
        n = len(vals)
        mean = sum(vals) / n
        sd = math.sqrt(sum((v - mean) ** 2 for v in vals) / (n - 1)) if n > 1 else 0.0
        prec = int(m.get("precision", _precision_of(vals)))
        tol = 0.5 * (10 ** (-prec)) if prec >= 0 else 0.5

        declared = m.get("mean")
        mean_ok = True
        if declared is not None:
            if abs(float(declared) - mean) > tol + 1e-12:
                mean_ok = False
                issues.append(
                    {
                        "level": "warn",
                        "where": f"measurements.{mid}",
                        "message": f"平均值不自洽：写了 {declared}，按 6 次算是 {round(mean, prec + 2)}（容差 {tol}）",
                    }
                )

        outliers = [v for v in vals if sd > 0 and abs(v - mean) > 3 * sd]
        for v in outliers:
            issues.append(
                {"level": "warn", "where": f"measurements.{mid}", "message": f"疑似离群值 {v}（>3σ）"}
            )
        if n >= 4 and sd == 0:
            issues.append({"level": "info", "where": f"measurements.{mid}", "message": "所有测量值完全相同，请确认不是抄错"})
        if n < 3:
            issues.append({"level": "info", "where": f"measurements.{mid}", "message": f"只有 {n} 次测量，A 类不确定度不可靠"})
        if not m.get("source"):
            issues.append({"level": "info", "where": f"measurements.{mid}", "message": "没有 source 定位"})

        scale = float(m.get("scale", 1.0))
        factor = UNITS.get(m.get("unit", ""), ("", 1.0))[1]
        base = mean * factor * scale
        names[mid] = base
        stats.append(
            {
                "id": mid,
                "label": m.get("label", ""),
                "unit": m.get("unit", ""),
                "n": n,
                "mean": mean,
                "meanInBase": base,
                "siFactor": factor * scale,
                "sd": sd,
                "min": min(vals),
                "max": max(vals),
                "declaredMean": declared,
                "precision": prec,
                "meanConsistent": mean_ok,
                "outliers": outliers,
            }
        )

    # ---------- 目标量重算
    targets = []
    for t in data.get("targets", []) or []:
        name = t.get("name") or t.get("expr", "?")
        expr = t.get("expr")
        if not expr:
            issues.append({"level": "error", "where": f"targets.{name}", "message": "缺 expr"})
            continue
        try:
            value = safe_eval(expr, names)
        except Exception as exc:  # noqa: BLE001
            issues.append({"level": "error", "where": f"targets.{name}", "message": f"表达式无法求值：{exc}"})
            continue
        names[name] = value
        item = {"name": name, "expr": expr, "value": value, "unit": t.get("unit", "")}
        expected = t.get("expected")
        if expected is not None:
            expected = float(expected)
            rel = abs(value - expected) / abs(expected) if expected else float("inf")
            tol = float(t.get("tolerance", 0.01))
            item["expected"] = expected
            item["relDiff"] = rel
            item["tolerance"] = tol
            if rel > tol:
                level = "error" if rel > 5 * tol else "warn"
                issues.append(
                    {
                        "level": level,
                        "where": f"targets.{name}",
                        "message": f"重算 {value:.6g} 与预期 {expected:.6g} 相差 {rel * 100:.2f}%（容差 {tol * 100:.2f}%）",
                    }
                )
            item["status"] = "一致" if rel <= tol else "不一致"
        else:
            item["status"] = "未给预期值"
        targets.append(item)

    errors = [i for i in issues if i["level"] == "error"]
    warns = [i for i in issues if i["level"] == "warn"]
    infos = [i for i in issues if i["level"] == "info"]

    return {
        "file": os.path.abspath(path) if path else "",
        "experiment": data.get("experiment", ""),
        "counts": {"quantities": len(quantities), "measurements": len(measurements), "targets": len(targets)},
        "quantities": q_report,
        "measurements": stats,
        "targets": targets,
        "issues": issues,
        "summary": {"error": len(errors), "warn": len(warns), "info": len(infos)},
        "ok": not errors,
    }


TEMPLATE = {
    "_说明": [
        "value + unit 按仪表/照片原样填，CLI 会自动按 unit 归一化到 SI（如 1116.12 g -> 1.11612）。",
        "targets.expr 里直接用归一化后的量名，单位保持 SI，不用自己写 /1000。",
        "measurements.scale 只在『记录的不是单个周期』时使用，例如记的是 50T 总时间则 scale=0.02。",
        "每个量/每张表都写 source（照片名 + 位置），校验失败时才能回看照片复核。",
    ],
    "experiment": "<实验名，如 扭摆法测量转动惯量>",
    "source": {"images": ["img1.jpg"], "note": "<照片说明>"},
    "quantities": {
        "m_sample": {"value": 0.0, "unit": "g", "uncertainty": 0.01, "source": "<照片+区域>"},
        "D_sample": {"value": 0.0, "unit": "mm", "source": "<照片+区域>"},
    },
    "measurements": [
        {
            "id": "T0",
            "label": "<这一组是什么，如 空盘>",
            "unit": "ms",
            "values": [0, 0, 0, 0, 0, 0],
            "mean": 0,
            "scale": 1,
            "precision": 0,
            "source": "<照片+行号>",
        }
    ],
    "targets": [
        {
            "name": "J1理",
            "expr": "1/8*m_sample*D_sample^2",
            "unit": "kg*m^2",
            "expected": 0.0,
            "tolerance": 0.01,
        }
    ],
}


def render(report: dict) -> str:
    out = []
    c = report["counts"]
    out.append(f"数据文件  : {report['file']}")
    out.append(f"实验      : {report['experiment'] or '(未填)'}")
    out.append(f"内容      : 量 {c['quantities']} 个 | 测量组 {c['measurements']} 组 | 重算目标 {c['targets']} 项")
    converted = [q for q in report.get("quantities", []) if q["factor"] != 1.0]
    if converted:
        out.append("单位归一化（表达式里用的 SI 值）:")
        for q in converted:
            out.append(f"  {q['key']:<10} {q['value']:g} {q['unit']}  ->  {q['valueSI']:.6g}")
    out.append("")
    if report["measurements"]:
        out.append("测量组统计:")
        for m in report["measurements"]:
            flag = "✓" if m["meanConsistent"] else "✗"
            factor = m.get("siFactor", 1.0)
            si = f"  = {m['meanInBase']:.6g}" if factor != 1.0 else ""
            out.append(
                f"  {m['id']:<6} {m['label'][:14]:<16} n={m['n']}  平均={m['mean']:.4g}{m['unit']}{si}"
                f"  标准差={m['sd']:.4g}  区间[{m['min']:.4g},{m['max']:.4g}]  平均值自洽 {flag}"
            )
    if report["targets"]:
        out.append("")
        out.append("重算目标:")
        for t in report["targets"]:
            extra = ""
            if "expected" in t:
                extra = f"  预期={t['expected']:.6g}  相对差={t['relDiff'] * 100:.3f}%  [{t['status']}]"
            out.append(f"  {t['name']:<10} = {t['value']:.6g} {t['unit']}{extra}")
    out.append("")
    order = {"error": 0, "warn": 1, "info": 2}
    for issue in sorted(report["issues"], key=lambda x: order[x["level"]]):
        out.append(f"  [{issue['level']:<5}] {issue['where']}: {issue['message']}")
    s = report["summary"]
    out.append("")
    out.append(
        f"结论: {'可以填报告' if report['ok'] else '有 error，必须回看照片复核'}"
        f"（error {s['error']} / warn {s['warn']} / info {s['info']}）"
    )
    return "\n".join(out)


def run(args) -> int:
    action = getattr(args, "data_cmd", None)

    if action == "template":
        text = json.dumps(TEMPLATE, ensure_ascii=False, indent=2)
        if getattr(args, "out", None):
            with open(args.out, "w", encoding="utf-8") as fh:
                fh.write(text)
            print(f"已写出模板：{args.out}")
        else:
            print(text)
        return 0

    if action != "check":
        print("用法：labreport data check <data.json> | labreport data template [-o data.json]")
        return 2

    path = args.data_file
    if not os.path.isfile(path):
        print(f"找不到数据文件：{path}")
        return 1
    try:
        with open(path, encoding="utf-8") as fh:
            data = json.load(fh)
    except Exception as exc:  # noqa: BLE001
        print(f"JSON 解析失败：{exc}")
        return 1

    report = check(data, path)
    if getattr(args, "json", False):
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(render(report))
    return 0 if report["ok"] else 1
