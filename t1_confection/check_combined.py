#!/usr/bin/env python3
"""Compatibility wrapper for :mod:`tools.analysis.check_combined`."""

from pathlib import Path
import runpy


_TARGET = Path(__file__).resolve().parents[1] / "tools" / "analysis" / "check_combined.py"
globals().update(runpy.run_path(str(_TARGET), run_name=__name__))
