#!/usr/bin/env python3
"""Compatibility wrapper for :mod:`tools.analysis.concat_all_scenarios`."""

from pathlib import Path
import runpy


_TARGET = Path(__file__).resolve().parents[1] / "tools" / "analysis" / "concat_all_scenarios.py"
globals().update(runpy.run_path(str(_TARGET), run_name=__name__))
