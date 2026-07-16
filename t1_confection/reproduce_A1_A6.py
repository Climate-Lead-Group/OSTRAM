#!/usr/bin/env python3
"""Compatibility wrapper for :mod:`tools.analysis.reproduce_A1_A6`."""

from pathlib import Path
import runpy


_TARGET = Path(__file__).resolve().parents[1] / "tools" / "analysis" / "reproduce_A1_A6.py"
globals().update(runpy.run_path(str(_TARGET), run_name=__name__))
