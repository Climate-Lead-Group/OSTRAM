"""Rule loaders refuse unknown configuration keys and read renamed values."""
from __future__ import annotations

import csv
import tempfile
import unittest
from pathlib import Path

import yaml

from ostram.pipeline.scenarios.rules import (
    build_cap_rule as lid_rule,
    add_storage_min_investment as storage,
    apply_base_year_pin as pin,
    apply_internal_transmission as internal_tx,
    apply_internal_tx_losses as internal_losses,
    relax_interconnectors as relax,
    set_interconnector_direction as direction,
    set_min_capacity_floors as floors,
    set_retirement_schedule as retirement,
    set_vre_targets as vre,
)

YAML_LOADERS = {
    "lid_rule": lid_rule.load_config,
    "storage_floors": storage.load_config,
    "relax_interconnectors": relax.load_config,
    "set_interconnector_direction": direction.load_config,
    "bau_calibration": floors.load_config,
    "retirement_schedule": retirement.load_config,
    "set_vre_targets": vre.load_config,
}


class RuleConfigKeyTests(unittest.TestCase):
    def setUp(self) -> None:
        self._temp = tempfile.TemporaryDirectory()
        self.root = Path(self._temp.name)

    def tearDown(self) -> None:
        self._temp.cleanup()

    def _write(self, name: str, data: dict) -> Path:
        path = self.root / name
        path.write_text(yaml.safe_dump(data), encoding="utf-8")
        return path

    def _assert_refused(self, loader, path: Path, key: str) -> None:
        with self.assertRaises(ValueError) as raised:
            loader(path)
        self.assertIn(str(path), str(raised.exception))
        self.assertIn(key, str(raised.exception))

    def test_rule_yaml_loaders_refuse_an_unknown_key(self) -> None:
        for name, loader in YAML_LOADERS.items():
            with self.subTest(loader=name):
                path = self._write(f"{name}.yaml", {"unknown_key": 1})
                self._assert_refused(loader, path, "unknown_key")

    def test_build_cap_rule_refuses_old_keys_and_names_the_new_one(self) -> None:
        for old, new in lid_rule.RENAMED_KEYS.items():
            with self.subTest(key=old):
                path = self._write("build_cap_rule.yaml", {old: 1})
                with self.assertRaises(ValueError) as raised:
                    lid_rule.load_config(path)
                self.assertIn(old, str(raised.exception))
                self.assertIn(new, str(raised.exception))

    def test_entry_with_unknown_key_is_refused(self) -> None:
        entry = {"cr": "BGDXX", "tech": "PWRSOL", "param": "min_cap",
                 "schedule": {2030: 1.0}, "unknown_entry_key": 1}
        path = self._write("bau_calibration.yaml", {"floors": [entry]})
        self._assert_refused(floors.load_config, path, "unknown_entry_key")

    def test_internal_transmission_block_refuses_an_unknown_key(self) -> None:
        path = self._write(
            "Config_country_codes.yaml",
            {"internal_transmission": {"transmission_loss": 0.03, "unknown_key": 1}},
        )
        for loader in (internal_tx.read_config, internal_losses.read_loss):
            with self.subTest(loader=loader.__module__):
                self._assert_refused(loader, path, "unknown_key")

    def test_relax_mode_cap_9999_and_its_previous_name(self) -> None:
        for name in ("cap_9999", "unconstrained"):
            with self.subTest(mode=name):
                path = self._write("relax_interconnectors.yaml", {"mode": name})
                self.assertEqual(relax.load_config(path)["mode"], "cap_9999")

    def test_pin_rules_refuse_an_unknown_column(self) -> None:
        path = self.root / "pin.csv"
        with path.open("w", encoding="utf-8", newline="") as stream:
            csv.writer(stream).writerow(pin.EXPECTED_FIELDS + ("unknown_column",))
        self._assert_refused(pin.load_pin_rules, path, "unknown_column")


if __name__ == "__main__":
    unittest.main()
