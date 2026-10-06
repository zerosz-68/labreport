"""支持 `python -m labreport ...`（未 pip 安装时也能直接用）。"""

from .cli import main

if __name__ == "__main__":
    raise SystemExit(main())
