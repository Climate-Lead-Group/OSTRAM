"""Prepare registered example bundles and generate profile-aware reports."""

from __future__ import annotations

import argparse
from pathlib import Path
import re
import shutil
from typing import Sequence

from ostram.paths import resolve_paths
from ostram.profile_workspace import prepare_profile
from ostram.profiles import PROFILE_ENV, PROFILE_MANIFEST_ENV, load_manifest
from ostram.reporting.training_dashboard import generate_report


_CAPTURE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_.-]*$")


def _active_manifest(profile_id: str):
    import os

    active = os.environ.get(PROFILE_ENV)
    if active != profile_id:
        raise RuntimeError(
            f"active profile is {active!r}, but command requested {profile_id!r}"
        )
    raw = os.environ.get(PROFILE_MANIFEST_ENV)
    if not raw:
        raise RuntimeError("profile manifest environment is missing")
    return load_manifest(Path(raw), expected_profile=profile_id)


def _result_candidates(paths, metadata) -> list[Path]:
    configured = metadata.get("reporting", {})
    pattern = configured.get("result_glob") if isinstance(configured, dict) else None
    if pattern:
        if Path(str(pattern)).is_absolute() or ".." in Path(str(pattern)).parts:
            raise ValueError("reporting.result_glob must be workspace-relative")
        candidates = list(paths.workspace.glob(str(pattern)))
    else:
        candidates = list(paths.execution_workspace.rglob("*Combined_Inputs_Outputs.csv"))
    return sorted(
        (path.resolve() for path in candidates if path.is_file()),
        key=lambda path: (path.stat().st_mtime_ns, str(path)),
    )


def _report(manifest, capture: str | None) -> Path:
    paths = resolve_paths()
    report_root = paths.workspace / "reports"
    snapshots_root = report_root / "snapshots"
    candidates = _result_candidates(paths, manifest.metadata)
    if capture is not None:
        if not _CAPTURE.fullmatch(capture):
            raise ValueError(f"unsafe capture label: {capture!r}")
        if not candidates:
            raise FileNotFoundError("no combined result CSV is available to capture")
        snapshots_root.mkdir(parents=True, exist_ok=True)
        destination = snapshots_root / f"{capture}.csv"
        if destination.exists():
            raise FileExistsError(f"capture already exists: {destination}")
        shutil.copy2(candidates[-1], destination)
    snapshots = [
        (path.stem, path)
        for path in sorted(snapshots_root.glob("*.csv"))
    ]
    if not snapshots and candidates:
        snapshots = [("current", candidates[-1])]
    if not snapshots:
        raise FileNotFoundError("no result snapshots are available for reporting")
    return generate_report(
        snapshots,
        report_root / f"{manifest.profile_id}.html",
        profile_id=manifest.profile_id,
        manifest=manifest.path,
        workspace=paths.workspace,
        metadata=manifest.metadata,
    )


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="python -m ostram example")
    subparsers = parser.add_subparsers(dest="action", required=True)
    prepare = subparsers.add_parser("prepare")
    prepare.add_argument("profile")
    prepare.add_argument("--reset", action="store_true")
    report = subparsers.add_parser("report")
    report.add_argument("profile")
    report.add_argument("--capture")
    args = parser.parse_args(argv)
    manifest = _active_manifest(args.profile)
    if args.action == "prepare":
        prepared = prepare_profile(manifest, paths=resolve_paths(), reset=args.reset)
        print(f"Prepared profile {manifest.profile_id}: {prepared.workspace}")
        return 0
    output = _report(manifest, args.capture)
    print(f"Profile report: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
