#!/usr/bin/env python3
"""Compatibility wrapper for :mod:`tools.analysis.ostram_trn_plotter`."""

from pathlib import Path
import runpy


_TARGET = Path(__file__).resolve().parents[1] / "tools" / "analysis" / "ostram_trn_plotter.py"
globals().update(runpy.run_path(str(_TARGET), run_name=__name__))
