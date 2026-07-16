#!/usr/bin/env python3
"""Compatibility wrapper for the interconnections analysis dashboard."""

from pathlib import Path
import runpy


_TARGET = (
    Path(__file__).resolve().parents[1]
    / "tools"
    / "analysis"
    / "visualization"
    / "Z_AUX_interconnections_dashboard.py"
)
globals().update(runpy.run_path(str(_TARGET), run_name=__name__))
