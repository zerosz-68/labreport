"""把适配页装进目标 agent（Claude / DSH / Codex / Cursor）。

设计：正文只有一份（`_body.md`），各 agent 的差别只在**外壳**（frontmatter / 文件位置）。
这样正文改一处、四份适配页同时更新，不会各写一份漂移。
"""

from __future__ import annotations

import os
import shutil

BODY = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_body.md")

MARK_BEGIN = "<!-- labreport:begin -->"
MARK_END = "<!-- labreport:end -->"

DESCRIPTION = (
    "大学实验报告填写与自检工作流（物理/化学等）：用户提供待填写的报告 Word 模板（.docx）、"
    "手写记录照片或数据表，要求'把数据填进报告''按模板填''检查报告填全没有'时使用。"
    "本 skill 是薄封装——读模板结构、校验数据、LaTeX 转 Word 原生公式、填写、自检、转 PDF "
    "全部交给已安装的 labreport 命令。"
)

AGENTS = {
    "claude": {
        "title": "Claude Code / Claude 桌面",
        "dest": "~/.claude/skills/labreport/SKILL.md",
        "wrap": lambda body: f"---\nname: labreport\ndescription: {DESCRIPTION}\n---\n\n{body}",
        "mode": "file",
    },
    "dsh": {
        "title": "DeepSeek Harness（DSH）",
        "dest": "$DSH_HOME/skills/labreport/SKILL.md",
        "wrap": lambda body: f"---\nname: labreport\ndescription: {DESCRIPTION}\n---\n\n{body}",
        "mode": "file",
    },
    "codex": {
        "title": "Codex CLI",
        "dest": "~/.codex/AGENTS.md",
        "wrap": lambda body: f"{MARK_BEGIN}\n## 实验报告（labreport）\n\n{body}\n{MARK_END}\n",
        "mode": "append",
    },
    "cursor": {
        "title": "Cursor",
        "dest": ".cursor/rules/labreport.mdc",
        "wrap": lambda body: (
            "---\n"
            f"description: {DESCRIPTION}\n"
            "globs: [\"*.docx\", \"**/*.docx\", \"**/data.json\", \"**/map.json\"]\n"
            "alwaysApply: false\n"
            "---\n\n"
            f"{body}"
        ),
        "mode": "file",
    },
}


def _expand(path: str) -> str:
    dsh_home = os.environ.get("DSH_HOME") or os.path.join(os.path.expanduser("~"), ".dsh")
    path = path.replace("$DSH_HOME", dsh_home)
    return os.path.abspath(os.path.expanduser(path))


def page_for(agent: str) -> str:
    with open(BODY, encoding="utf-8") as fh:
        body = fh.read()
    return AGENTS[agent]["wrap"](body)


def _write_append(path: str, content: str) -> str:
    """AGENTS.md 这类共享文件：用标记块替换，避免重复堆积。"""
    old = ""
    if os.path.isfile(path):
        with open(path, encoding="utf-8") as fh:
            old = fh.read()
    if MARK_BEGIN in old and MARK_END in old:
        start = old.index(MARK_BEGIN)
        end = old.index(MARK_END) + len(MARK_END)
        new = old[:start] + content.strip() + old[end:]
        action = "已替换标记块"
    else:
        new = (old.rstrip() + "\n\n" if old.strip() else "") + content
        action = "已追加标记块"
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(new)
    return action


def run(args) -> int:
    if getattr(args, "list", False) or not getattr(args, "for_agent", None):
        print("支持的 agent 与默认落点：")
        for key, cfg in AGENTS.items():
            print(f"  {key:<8} {cfg['title']:<24} {cfg['dest']}")
        print("\n用法：labreport install-skill --for <agent> [--dest 路径] [--print|--dry-run|--force]")
        return 0

    agent = args.for_agent
    if agent not in AGENTS:
        print(f"不支持的 agent：{agent}（可选：{', '.join(AGENTS)}）")
        return 1

    cfg = AGENTS[agent]
    content = page_for(agent)

    if getattr(args, "print_page", False):
        print(content)
        return 0

    dest = _expand(args.dest or cfg["dest"])
    if getattr(args, "dry_run", False):
        print(f"[dry-run] 目标：{dest}")
        print(f"[dry-run] 模式：{cfg['mode']}    内容 {len(content)} 字符")
        print("-" * 60)
        print(content[:400] + ("\n…" if len(content) > 400 else ""))
        return 0

    try:
        if cfg["mode"] == "append":
            action = _write_append(dest, content)
        else:
            if os.path.isfile(dest) and not getattr(args, "force", False):
                with open(dest, encoding="utf-8") as fh:
                    if fh.read() == content:
                        print(f"已是最新，无需改动：{dest}")
                        return 0
                print(f"目标已存在且内容不同：{dest}\n加 --force 覆盖，或先用 --print 看内容。")
                return 1
            os.makedirs(os.path.dirname(dest) or ".", exist_ok=True)
            with open(dest, "w", encoding="utf-8") as fh:
                fh.write(content)
            action = "已写入"
    except Exception as exc:  # noqa: BLE001
        print(f"写入失败：{exc}")
        return 1

    print(f"{action}：{dest}")
    print(f"（{cfg['title']}）现在可以在该 agent 里直接说「按模板填这份实验报告」了。")
    if agent == "codex":
        print("提示：AGENTS.md 是共享文件，本 skill 放在标记块内，重复安装只替换标记块。")
    return 0
