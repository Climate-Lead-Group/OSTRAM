#!/usr/bin/env python3
"""Compatibility wrapper for the transmission analysis map generator."""

from pathlib import Path
import runpy


_TARGET = (
    Path(__file__).resolve().parents[1]
    / "tools"
    / "analysis"
    / "visualization"
    / "Z_AUX_generate_transmission_maps.py"
)
globals().update(runpy.run_path(str(_TARGET), run_name=__name__))
