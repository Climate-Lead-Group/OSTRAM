"""Strict YAML reading for repository and user files.

``yaml.safe_load`` keeps the last of two equal keys in one mapping and raises
nothing, so a repeated block silently replaces the first. The reader here
refuses a repeated key at any depth and otherwise returns exactly what
``yaml.safe_load`` returns.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import yaml

_MERGE_TAG = "tag:yaml.org,2002:merge"
_PROJECT_ROOT = Path(__file__).resolve().parents[1]


class DuplicateKeyError(ValueError):
    """Raised when one YAML mapping holds the same key twice."""


def display_path(path: Any) -> str:
    """Return ``path`` relative to the repository root when it lies inside it."""
    try:
        return Path(path).resolve().relative_to(_PROJECT_ROOT).as_posix()
    except (TypeError, ValueError, OSError):
        return str(path)


class StrictLoader(yaml.SafeLoader):
    """``yaml.SafeLoader`` that raises DuplicateKeyError for a repeated mapping key."""

    source = "<string>"

    def construct_mapping(self, node, deep=False):
        if isinstance(node, yaml.MappingNode):
            first_lines: dict[Any, int] = {}
            for key_node, _ in node.value:
                if key_node.tag == _MERGE_TAG:
                    continue
                key = self.construct_object(key_node, deep=True)
                line = key_node.start_mark.line + 1
                try:
                    seen = key in first_lines
                except TypeError:
                    continue  # an unhashable key; SafeLoader reports it below
                if seen:
                    raise DuplicateKeyError(
                        f"{self.source}, line {line}: the key {key!r} appears a second "
                        f"time (first at line {first_lines[key]}). Keep one."
                    )
                first_lines[key] = line
        return super().construct_mapping(node, deep=deep)


def safe_load(stream: Any, source: Any = None) -> Any:
    """Return what ``yaml.safe_load(stream)`` returns, or raise DuplicateKeyError.

    ``source`` names the file in the error; it defaults to the stream's name.
    """
    loader = StrictLoader(stream)
    if source is None:
        source = getattr(stream, "name", None)
    if source is not None:
        loader.source = display_path(source)
    try:
        return loader.get_single_data()
    finally:
        loader.dispose()


def load_file(path: Any) -> Any:
    """Read the UTF-8 YAML file at ``path`` with the strict reader."""
    path = Path(path)
    return safe_load(path.read_text(encoding="utf-8"), source=path)
