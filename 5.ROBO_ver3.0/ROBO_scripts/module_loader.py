"""Utility for dynamically loading step helper modules with dotted filenames."""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path
from types import ModuleType

_CACHE: dict[str, ModuleType] = {}


def load_helper(module_name: str) -> ModuleType:
    """Load a helper module (e.g., "Ac.KANRI_spreadsheet") and cache it."""

    if module_name in _CACHE:
        return _CACHE[module_name]

    base_dir = Path(__file__).resolve().parent
    candidate_paths = [
        base_dir / f"{module_name}.py",
        base_dir / Path(module_name.replace('.', '/')).with_suffix('.py'),
        base_dir / module_name / f"{module_name}.py",
        base_dir / module_name / "__init__.py",
    ]

    file_path = None
    for candidate in candidate_paths:
        if candidate and candidate.exists():
            file_path = candidate
            break

    if file_path is None:
        fallback = list(base_dir.rglob(f"{module_name}.py"))
        if fallback:
            file_path = fallback[0]

    if file_path is None:
        raise FileNotFoundError(
            f"モジュールファイルが見つかりません: {candidate_paths[0]}"
        )

    spec_name = f"chouji_{module_name.replace('.', '_')}"
    spec = importlib.util.spec_from_file_location(spec_name, file_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"モジュール {module_name} をロードできませんでした。")

    module = importlib.util.module_from_spec(spec)
    sys.modules[spec_name] = module
    spec.loader.exec_module(module)
    _CACHE[module_name] = module
    return module
