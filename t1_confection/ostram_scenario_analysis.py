#!/usr/bin/env python3
"""Compatibility wrapper for :mod:`tools.analysis.ostram_scenario_analysis`."""

from pathlib import Path
import runpy


_TARGET = Path(__file__).resolve().parents[1] / "tools" / "analysis" / "ostram_scenario_analysis.py"
globals().update(runpy.run_path(str(_TARGET), run_name=__name__))
