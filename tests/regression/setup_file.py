"""The setup file of the default setup, from the path resolver."""

from pathlib import Path

from ostram.paths import resolve_paths

SETUP_FILE = resolve_paths(
    project_root=Path(__file__).resolve().parents[2], environ={}
).scenario_registry
