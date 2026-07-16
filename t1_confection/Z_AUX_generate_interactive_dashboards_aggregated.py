#!/usr/bin/env python3
"""Compatibility wrapper for the aggregated analysis dashboard generator."""

from pathlib import Path
import runpy


_TARGET = (
    Path(__file__).resolve().parents[1]
    / "tools"
    / "analysis"
    / "visualization"
    / "Z_AUX_generate_interactive_dashboards_aggregated.py"
)
globals().update(runpy.run_path(str(_TARGET), run_name=__name__))
