"""Stage 11 file-module loader; removed after Stage 12 package migration."""

from __future__ import annotations

import importlib.util
from pathlib import Path
import sys


def load_file_module(name: str, path: str | Path):
    module_path = Path(path).resolve()
    existing = sys.modules.get(name)
    if existing is not None:
        return existing
    spec = importlib.util.spec_from_file_location(name, module_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"cannot load Python module from {module_path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    try:
        spec.loader.exec_module(module)
    except Exception:
        sys.modules.pop(name, None)
        raise
    return module
