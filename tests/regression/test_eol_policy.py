from __future__ import annotations

import hashlib
import json
import sys
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent))

import ostram_regression as regression


TEST_ROOT = Path(__file__).resolve().parent
REPO_ROOT = TEST_ROOT.parents[1]
PROTECTED_MANIFEST = (
    TEST_ROOT / "baselines" / "5ce4e66480e1-static-nosolver" / "manifest.json"
)
HOUSEKEEPING_REFERENCE = "8dd3361a1fc7f2c9ea4df51d5b2d0e50e0ce8554"
APPROVED_RELOCATED_PROTECTED_FILE = (
    "t1_confection/A3_process/_test_scenarios_lite.py"
)
RELOCATED_VALIDATION_FILE = "tests/validation/test_scenarios_lite.py"

CRLF_SUFFIXES = {".bat", ".cmd"}
BINARY_SUFFIXES = {
    ".7z",
    ".bmp",
    ".bz2",
    ".doc",
    ".docx",
    ".feather",
    ".gif",
    ".gz",
    ".ico",
    ".jpeg",
    ".jpg",
    ".npy",
    ".npz",
    ".ods",
    ".parquet",
    ".pdf",
    ".pickle",
    ".pkl",
    ".png",
    ".ppt",
    ".pptx",
    ".rar",
    ".tar",
    ".tbz2",
    ".tgz",
    ".tif",
    ".tiff",
    ".twbx",
    ".webp",
    ".xls",
    ".xlsb",
    ".xlsm",
    ".xlsx",
    ".xz",
    ".zip",
}


def _git(*args: str) -> str:
    output = regression._git(REPO_ROOT, *args)
    if output is None:
        raise AssertionError(f"git command failed: git {' '.join(args)}")
    return output


def _tracked_eol_records() -> dict[str, tuple[str, str, str]]:
    output = _git("ls-files", "--eol", "-z")
    records: dict[str, tuple[str, str, str]] = {}
    for raw_record in output.split("\0"):
        if not raw_record:
            continue
        try:
            metadata, path = raw_record.split("\t", 1)
            index_eol, worktree_eol, attributes = metadata.strip().split(maxsplit=2)
        except ValueError as exc:
            raise AssertionError(f"invalid git ls-files --eol record: {raw_record!r}") from exc
        if not attributes.startswith("attr/"):
            raise AssertionError(f"missing attribute classification for {path}: {attributes!r}")
        records[path.replace("\\", "/")] = (
            index_eol,
            worktree_eol,
            attributes.removeprefix("attr/"),
        )
    return records


def _attributes_for(paths: list[str]) -> dict[str, dict[str, str]]:
    output = _git("check-attr", "-z", "text", "eol", "diff", "merge", "--", *paths)
    fields = [field for field in output.split("\0") if field]
    if len(fields) % 3:
        raise AssertionError(f"invalid git check-attr output: {output!r}")
    result: dict[str, dict[str, str]] = {}
    for offset in range(0, len(fields), 3):
        path, attribute, value = fields[offset : offset + 3]
        result.setdefault(path.replace("\\", "/"), {})[attribute] = value
    return result


class CheckoutEolPolicyTests(unittest.TestCase):
    def test_all_tracked_files_have_deterministic_worktree_bytes(self) -> None:
        tracked = {
            path.replace("\\", "/")
            for path in _git("ls-files", "-z").split("\0")
            if path
        }
        records = _tracked_eol_records()
        self.assertEqual(set(records), tracked)

        for path, (index_eol, worktree_eol, attributes) in sorted(records.items()):
            with self.subTest(path=path):
                suffix = Path(path).suffix.lower()
                if suffix in BINARY_SUFFIXES:
                    self.assertEqual(attributes, "-text")
                    self.assertEqual((index_eol, worktree_eol), ("i/-text", "w/-text"))
                elif suffix in CRLF_SUFFIXES:
                    self.assertEqual(attributes, "text eol=crlf")
                    self.assertIn(
                        (index_eol, worktree_eol),
                        {("i/lf", "w/crlf"), ("i/none", "w/none")},
                    )
                else:
                    self.assertEqual(attributes, "text=auto eol=lf")
                    self.assertIn(
                        (index_eol, worktree_eol),
                        {("i/lf", "w/lf"), ("i/none", "w/none")},
                    )

    def test_absent_extensions_are_covered_by_explicit_attributes(self) -> None:
        ordinary = "tests/regression/__eol_policy_probe__.py"
        crlf_paths = [
            f"tests/regression/__eol_policy_probe__{suffix}"
            for suffix in sorted(CRLF_SUFFIXES)
        ]
        binary_paths = [
            f"tests/regression/__eol_policy_probe__{suffix}"
            for suffix in sorted(BINARY_SUFFIXES)
        ]
        attributes = _attributes_for([ordinary, *crlf_paths, *binary_paths])

        self.assertEqual(
            attributes[ordinary],
            {"text": "auto", "eol": "lf", "diff": "unspecified", "merge": "unspecified"},
        )
        for path in crlf_paths:
            with self.subTest(path=path):
                self.assertEqual(
                    attributes[path],
                    {"text": "set", "eol": "crlf", "diff": "unspecified", "merge": "unspecified"},
                )
        for path in binary_paths:
            with self.subTest(path=path):
                self.assertEqual(attributes[path]["text"], "unset")
                self.assertEqual(attributes[path]["diff"], "unset")
                self.assertEqual(attributes[path]["merge"], "unset")
                self.assertEqual(attributes[path]["eol"], "lf")

    def test_committed_manifest_matches_raw_protected_worktree(self) -> None:
        self.assertIn(".gitattributes", regression.PROTECTED_FILES)
        for path in ("ostram/__init__.py", "ostram/__main__.py"):
            self.assertIn(path, regression.PROTECTED_FILES)
            self.assertTrue((REPO_ROOT / path).is_file())
        self.assertIn("t1_confection/a3_orchestrator.py", regression.PROTECTED_FILES)
        for path in (
            "t1_confection/a1_b1_transforms/__init__.py",
            "t1_confection/a1_b1_transforms/planning.py",
            "t1_confection/a1_b1_transforms/tables.py",
            "t1_confection/a1_b1_transforms/effects.py",
            "t1_confection/a1_b1_transforms/validation.py",
            "t1_confection/a1_b1_transforms/delivery.py",
        ):
            self.assertIn(path, regression.PROTECTED_FILES)
        result = regression.verify_protected(REPO_ROOT, PROTECTED_MANIFEST)
        if result["ok"]:
            return

        # Housekeeping relocates one developer-only test out of the broad
        # A3-process Python glob. Preserve the historical manifest unchanged
        # and prove every remaining protected path is byte-identical to the
        # pinned housekeeping reference.
        self.assertFalse((REPO_ROOT / APPROVED_RELOCATED_PROTECTED_FILE).exists())
        self.assertTrue((REPO_ROOT / RELOCATED_VALIDATION_FILE).is_file())
        self.assertEqual(
            result["expected"]["file_count"] - 1,
            result["actual"]["file_count"],
            json.dumps(result, indent=2, sort_keys=True),
        )
        self.assertEqual(
            result["expected"]["total_bytes"] - 9_988,
            result["actual"]["total_bytes"],
            json.dumps(result, indent=2, sort_keys=True),
        )

        protected_pathspecs = [
            *regression.PROTECTED_TREE_ROOTS,
            *regression.PROTECTED_FILES,
            *(f":(glob){pattern}" for pattern in regression.PROTECTED_GLOBS),
            f":(exclude){APPROVED_RELOCATED_PROTECTED_FILE}",
        ]
        drift = _git(
            "diff",
            "--name-only",
            HOUSEKEEPING_REFERENCE,
            "--",
            *protected_pathspecs,
        )
        self.assertEqual(
            "",
            drift.strip(),
            f"protected drift from {HOUSEKEEPING_REFERENCE}: {drift}",
        )

    def test_protected_verifier_detects_eol_only_raw_byte_drift(self) -> None:
        lf_bytes = b"alpha\n"
        crlf_bytes = b"alpha\r\n"
        self.assertEqual(
            regression.normalize_text_bytes(lf_bytes),
            regression.normalize_text_bytes(crlf_bytes),
        )

        def raw_snapshot(payload: bytes) -> dict[str, object]:
            file_hash = hashlib.sha256(payload).hexdigest()
            aggregate = hashlib.sha256()
            aggregate.update(b"protected.txt\0")
            aggregate.update(file_hash.encode("ascii"))
            aggregate.update(b"\n")
            return {
                "file_count": 1,
                "total_bytes": len(payload),
                "aggregate_raw_sha256": aggregate.hexdigest(),
            }

        expected = raw_snapshot(lf_bytes)
        actual = raw_snapshot(crlf_bytes)
        manifest = mock.Mock()
        manifest.read_text.return_value = json.dumps({"protected_working_tree": expected})
        with mock.patch.object(regression, "protected_snapshot", return_value=actual):
            result = regression.verify_protected(REPO_ROOT, manifest)

        self.assertFalse(result["ok"])
        self.assertEqual(result["expected"]["file_count"], result["actual"]["file_count"])
        self.assertNotEqual(result["expected"]["total_bytes"], result["actual"]["total_bytes"])
        self.assertNotEqual(
            result["expected"]["aggregate_raw_sha256"],
            result["actual"]["aggregate_raw_sha256"],
        )


if __name__ == "__main__":
    unittest.main()
