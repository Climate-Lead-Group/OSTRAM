"""Imports are not counted in the reserve-margin requirement.

``DSPTRN<node>`` mode 2 delivers imports onto ``ELC<node>01``, the fuel that
``ReserveMarginTagFuel`` tags. The left side of ``RM3_ReserveMargin_Constraint``
therefore leaves out every ``DSPTRN`` technology, so the requirement covers
plant output only. The storage-delay copy of the model file keeps the same line.
"""
from __future__ import annotations

import unittest
from pathlib import Path

from ostram.pipeline.execution.model_file_adjustments import storage_delay


MODEL = Path(__file__).resolve().parents[2] / "model" / "osemosys_fast_preprocessed.txt"
CONDITION = '(m,t) in MODExTECHNOLOGYperFUELout[f]: substr(t,1,6) != "DSPTRN"}'


def _rm3(lines):
    found = [l for l in lines if l.startswith("s.t. RM3_ReserveMargin_Constraint")]
    assert len(found) == 1, found
    return found[0]


class ReserveMarginImportTests(unittest.TestCase):
    def setUp(self) -> None:
        self.lines = MODEL.read_text(encoding="utf-8").splitlines()

    def test_left_side_leaves_out_dsptrn(self) -> None:
        left, right = _rm3(self.lines).split("<=")
        self.assertIn(CONDITION, left)
        self.assertIn("ReserveMarginTagFuel", left)
        self.assertNotIn("DSPTRN", right)

    def test_comment_above_states_the_reason(self) -> None:
        i = self.lines.index(_rm3(self.lines))
        self.assertTrue(self.lines[i - 1].startswith("# Imports are left out of the reserve requirement"))

    def test_storage_delay_copy_keeps_the_condition(self) -> None:
        self.assertEqual(_rm3(storage_delay.adjust_model_lines(self.lines)), _rm3(self.lines))


if __name__ == "__main__":
    unittest.main()
