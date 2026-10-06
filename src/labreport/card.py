"""实验知识卡：把"每个实验特有的专业内容"从流程里拆出来。

- 包内预置卡（`labreport/cards/*.md`）：随 CLI 分发，人人可用；
- 用户卡（默认 `~/.labreport/cards/*.md`，可用 `LABREPORT_CARD_DIR` 覆盖）：同名覆盖包内卡。

这样"通用流程"与"实验知识"彻底分离：换实验只需要一张卡，不用改 CLI。
自建卡目录是**每次调用现读环境变量**的（惰性），所以在同一进程里改配置也立即生效。
"""

from __future__ import annotations

import os
import shutil

PACKAGE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cards")
TEMPLATE = "_template.md"


def user_dir() -> str:
    """自建卡目录（惰性读取，便于进程内改配置与测试）。"""
    return os.environ.get("LABREPORT_CARD_DIR") or os.path.expanduser(
        os.path.join("~", ".labreport", "cards")
    )


def _norm(name: str) -> str:
    name = os.path.basename(name.strip())
    return name if name.endswith(".md") else name + ".md"


def list_cards() -> dict:
    """返回 {卡名: (路径, 来源)}；用户卡覆盖同名包内卡。"""
    out = {}
    if os.path.isdir(PACKAGE_DIR):
        for f in sorted(os.listdir(PACKAGE_DIR)):
            if f.endswith(".md") and f != TEMPLATE:
                out[f[:-3]] = (os.path.join(PACKAGE_DIR, f), "内置")
    udir = user_dir()
    if os.path.isdir(udir):
        for f in sorted(os.listdir(udir)):
            if f.endswith(".md"):
                out[f[:-3]] = (os.path.join(udir, f), "自建")
    return out


def _summary(path: str) -> str:
    try:
        with open(path, encoding="utf-8") as fh:
            for line in fh:
                s = line.strip()
                if s and not s.startswith("#") and not s.startswith(">"):
                    return s[:60]
    except OSError:
        pass
    return ""


def run(args) -> int:
    action = getattr(args, "card_cmd", None)
    udir = user_dir()

    if action == "list" or not action:
        cards = list_cards()
        if not cards:
            print("还没有任何知识卡。用 `labreport card new <实验名>` 建一张。")
            return 0
        print(f"实验知识卡（{len(cards)} 张）：")
        for name, (path, src) in cards.items():
            print(f"  {name:<26} [{src}] {_summary(path)}")
        print(f"\n包内卡目录：{PACKAGE_DIR}")
        print(f"自建卡目录：{udir}（同名会覆盖包内卡；可用 LABREPORT_CARD_DIR 改）")
        print("用 `labreport card show <卡名>` 查看内容。")
        return 0

    if action == "show":
        name = _norm(args.name)
        path = None
        if os.path.isfile(os.path.join(udir, name)):
            path = os.path.join(udir, name)
        elif os.path.isfile(os.path.join(PACKAGE_DIR, name)):
            path = os.path.join(PACKAGE_DIR, name)
        if not path:
            cards = list_cards()
            print(f"没有这张卡：{args.name}")
            print("可用：" + ("、".join(cards) if cards else "（暂无）"))
            return 1
        with open(path, encoding="utf-8") as fh:
            print(fh.read())
        return 0

    if action == "new":
        name = _norm(args.name)
        dest = os.path.abspath(os.path.expanduser(args.out or os.path.join(udir, name)))
        if os.path.isfile(dest) and not getattr(args, "force", False):
            print(f"已存在：{dest}\n加 --force 覆盖，或用 `labreport card show` 看现有内容。")
            return 1
        try:
            os.makedirs(os.path.dirname(dest) or ".", exist_ok=True)
            shutil.copyfile(os.path.join(PACKAGE_DIR, TEMPLATE), dest)
        except OSError as exc:
            print(f"写入失败：{exc}")
            return 1
        print(f"已从模板创建：{dest}")
        print("填完后再跑一次 `labreport card list` 就能看到它；同名时它优先于内置卡。")
        return 0

    print("用法：labreport card list | labreport card show <卡名> | labreport card new <卡名> [-o 路径]")
    return 2
