"""Import-safe helpers for the A1/A2-to-B1 transformation boundary.

The ``B1_Compiler.py`` entrypoint is responsible for orchestration.
Modules in this package perform no filesystem or pipeline work at import time.
"""

from __future__ import annotations


__all__ = ["delivery", "effects", "planning", "tables", "validation"]
