"""Exercise the TRN template choice without executing the file's pipeline driver."""
import ast
from pathlib import Path
import unittest

SOURCE = Path(__file__).resolve().parents[2] / "ostram/pipeline/scenarios/transformations/extract_ao_extensions.py"
NAMES = {"parse_code", "pick_trn_template"}
namespace = {}
tree = ast.parse(SOURCE.read_text(encoding="utf-8"))
exec(compile(ast.Module(body=[n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name in NAMES], type_ignores=[]), str(SOURCE), "exec"), namespace)


class TrnTemplateChoiceTests(unittest.TestCase):
    def test_tie_breaks_on_whole_name(self):
        # Both candidates share the first character and every other key.
        cands = ["TRNINDEAINDNE", "TRNBTNXXINDNE"]
        for order in (cands, list(reversed(cands))):
            self.assertEqual(
                namespace["pick_trn_template"](order, "BGDXX", "INDNE"),
                "TRNBTNXXINDNE",
            )

    def test_cross_border_ranks_before_name(self):
        cands = ["TRNABCDEINDNE", "TRNINDEAINDNE"]
        self.assertEqual(
            namespace["pick_trn_template"](cands, "BGDXX", "INDNE"),
            "TRNINDEAINDNE",
        )


if __name__ == "__main__":
    unittest.main()
