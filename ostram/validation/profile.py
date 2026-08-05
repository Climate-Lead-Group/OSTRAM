"""Profile-level validation and reportable effective values."""

from __future__ import annotations

from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

from ostram.paths import ProjectPaths
from ostram.profiles import ProfileManifest


def validate_profile(
    manifest: ProfileManifest,
    *,
    paths: ProjectPaths,
) -> dict[str, Any]:
    authorities = manifest.source_paths(paths, require_exists=True)
    effective = dict(manifest.metadata.get("effective_values", {}))
    for name, value in effective.items():
        try:
            number = Decimal(str(value))
        except (InvalidOperation, ValueError) as error:
            raise ValueError(f"effective value {name!r} is not numeric: {value!r}") from error
        if not number.is_finite():
            raise ValueError(f"effective value {name!r} is not finite")
    return {
        "profile_id": manifest.profile_id,
        "manifest": str(manifest.path),
        "manifest_sha256": manifest.digest,
        "authorities": {name: str(Path(path)) for name, path in authorities.items()},
        "policies": dict(manifest.policies),
        "effective_values": effective,
    }
