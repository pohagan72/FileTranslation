"""Make `filetranslation_mcp` importable from `src/` without an editable install.

Lets `pytest mcp_server/tests` work out-of-the-box; once the package is
`pip install -e .`'d (per pyproject.toml), this becomes a no-op.
"""

from __future__ import annotations

import sys
from pathlib import Path

_SRC = Path(__file__).resolve().parents[1] / "src"
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))
