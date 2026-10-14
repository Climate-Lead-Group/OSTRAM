"""conda_exe() finds conda on a standard Windows install and with a separate env folder on PATH."""

from __future__ import annotations

import os
from pathlib import Path
import tempfile
import unittest
from unittest import mock

from ostram.pipeline import orchestration


def _stub(path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("stub\n", encoding="utf-8")
    path.chmod(0o755)
    return path


def _conda_exe_with_path(*folders: Path) -> str:
    environment = {
        key: value for key, value in os.environ.items() if key.upper() != "CONDA_EXE"
    }
    environment["PATH"] = os.pathsep.join(str(folder) for folder in folders)
    with mock.patch.dict(os.environ, environment, clear=True):
        return orchestration.conda_exe()


class CondaExeTests(unittest.TestCase):
    def test_standard_install_with_only_condabin_on_path(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp) / "anaconda3"
            _stub(root / "condabin" / "conda.bat")
            conda = _stub(root / "Scripts" / "conda.exe")
            found = _conda_exe_with_path(root / "condabin")
        self.assertEqual(os.path.normcase(found), os.path.normcase(str(conda)))

    def test_r7_layout_resolves_to_scripts_conda_exe(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp) / "anaconda3"
            env = Path(temp) / "conda-envs" / "ostram-reproduce"
            _stub(env / "python.exe")
            _stub(root / "condabin" / "conda.bat")
            conda = _stub(root / "Scripts" / "conda.exe")
            found = _conda_exe_with_path(
                env, env / "Scripts", root / "Scripts", root / "condabin"
            )
        self.assertEqual(os.path.normcase(found), os.path.normcase(str(conda)))


if __name__ == "__main__":
    unittest.main()
