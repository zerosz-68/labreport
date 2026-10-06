#!/usr/bin/env python
"""零安装入口：不装包也能直接跑。

    python run.py doctor
    python run.py inspect <模板.docx>
    python run.py omml <公式.md> --compare <对照片段目录>

正式使用建议 `pip install -e .`（或 pip install git+...），那样任何目录下都有
`labreport` 命令；这个脚本只是给"不方便装东西"的环境兜底。
"""

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "src"))

from labreport.cli import main  # noqa: E402

if __name__ == "__main__":
    raise SystemExit(main())
