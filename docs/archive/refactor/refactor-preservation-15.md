# Refactor preservation validation: 15 decision scenarios

> **Historical checkpoint:** This record preserves the no-solver evidence available
> at commit `87c3ce7`. Current authority is the portable
> [accepted compiled-input baseline](../../../tests/regression/reports/accepted_compiled_solver_baseline_15.json)
> derived from the protected final manifest.
>
## Conclusion

**PASS.** Commit `87c3ce7cbfe04015f7b3e976d5838aa7e85165eb` preserves the accepted pre-refactor behavior and generated no-solver inputs for the 15 decision-relevant scenarios. The candidate produced all 15 canonical solver-consumed TXT files byte-for-byte and normalized-exact against the accepted reference, with zero parameter or set drift and zero duplicate declarations, keys, or memberships.

No solver, matrix, solution, result, DVC, batch, MOSOX, or unrestricted top-level pipeline execution occurred at this checkpoint. Solver-output equivalence was then intentionally deferred to `validation/solver-baseline-15`; the later accepted final baseline is identified above without changing this checkpoint's claim.

## Candidate and tree identity

- Branch: `validation/refactor-preservation-15`
- Candidate commit: `87c3ce7cbfe04015f7b3e976d5838aa7e85165eb`
- `origin/main`: `87c3ce7cbfe04015f7b3e976d5838aa7e85165eb`
- Merged CLI commit: `37516bca06026286489a069ea1607290fcfca4d4`
- Candidate, merged CLI, and current Git tree: `10d425bbee31863fff6df9f0a0e2a2f189dc5bff`
- Starting worktree: clean
- `.agents/`: present and untouched

The CLI commit is an ancestor of the candidate, and the relevant Git trees for `37516bc`, `87c3ce7`, and the validation start were identical.

## Accepted reference lineage

The directory `tests/regression/baselines/5ce4e66480e1-static-nosolver` is authoritative for the static pre-refactor gate, but is not alone proof of final compiled-input preservation.

- Pre-refactor baseline commit: `5ce4e66480e1326796406915885e972ef3687f5f`
- Pre-refactor tree: `234bbaca96075e8d2945cfd8f955e387ba4121fa`
- Baseline harness commit: `3cdd9fad0d2c5c4503a6f50bcf47ad63595c36cb`
- Baseline manifest SHA-256: `d9b67ad13c44f0c654accfc0c9aa38192bf1489fe4c14746a9b997c68c066a6d`
- Baseline hashes inventory SHA-256: `80577ad70ca0bd5d6ff286a9b203d0446eb3e5ced94546c2fa122bf4bbddd036`
- Baseline cleanup acceptance SHA-256: `8263d54c60e3a3cac57210289d0cff27ced2afa88e6095af3075401db01827a0`

The accepted final-input lineage is:

- Verified source commit: `41a54e51fd5a0776569b4900c44c624f09cc1f09`
- Source tree: `54f9e174fde266013ee9a20ff89f618140077c26`
- Committed pre-correction equivalence report:
  `tests/regression/reports/pre_correction_41a54e5_compiled_input_equivalence_15.json`
- Report SHA-256: `85c8489e65e8028dd5955dbbf204f2222796799d132ba979cf238e60008b8286`
- Read-only reference checkout: `OSTRAM_mainredo`, branch `ws3-phaseb-cleanredo`, commit `f1db168c8db0b61d03898f68cef2a7f28eccc80a`, tree `800103e866c4ce7f119f5887bdfd037bd46a310b`
- External B1 comparison report SHA-256: `c9773ac0dbf6e983ad5aa712b46144af078fc8d1a07bdd5ac6b1a76bd807dcb3`
- External B1 generation report SHA-256: `993a95605853e9b668e57bf7feb3f89e452d02903d4be40d05571fcc81b4e142`

The generated final files in `OSTRAM_mainredo` are ignored/untracked working files. Their sizes and SHA-256 values were independently matched to the committed 15-file report before they were used as the byte reference. This committed lineage, rather than directory presence alone, establishes the accepted reference.

## Scenario policy

The exact 20 protected scenarios are:

1. `BAU`
2. `A_Calibrated_BAU`
3. `A_Calibrated_BAU_Clipped`
4. `B_Optimised_VRE`
5. `B_Opt_Clipped`
6. `B_Opt_DirBidir`
7. `B_Opt_DirContractual`
8. `B_Opt_IndiaCosts`
9. `B_Opt_IndiaCostsFuel`
10. `B_Opt_LinkFreeze`
11. `B_Opt_SolarCapex130`
12. `B_Opt_SolarCapexHi`
13. `B_Opt_SolarCapexSpike`
14. `B_Opt_SolarHi10`
15. `B_Opt_TradeCap15`
16. `B_Opt_TradeCap30`
17. `B_Opt_TradeCap50`
18. `B_Opt_TxCap150`
19. `C_Target_VRE`
20. `C_Target_VRE_Clipped`

The exact 16 static cleanup-acceptance scenarios are:

1. `BAU`
2. `A_Calibrated_BAU`
3. `A_Calibrated_BAU_Clipped`
4. `B_Optimised_VRE`
5. `B_Opt_Clipped`
6. `B_Opt_DirBidir`
7. `B_Opt_DirContractual`
8. `B_Opt_IndiaCosts`
9. `B_Opt_IndiaCostsFuel`
10. `B_Opt_SolarCapex130`
11. `B_Opt_SolarCapexHi`
12. `B_Opt_SolarCapexSpike`
13. `B_Opt_TradeCap15`
14. `B_Opt_TxCap150`
15. `C_Target_VRE`
16. `C_Target_VRE_Clipped`

The 15 decision scenarios are the 16-scenario list excluding plain `BAU`. `A_Calibrated_BAU` remains the decision baseline. Plain `BAU` remains protected. The four protected superseded scenarios are `B_Opt_LinkFreeze`, `B_Opt_SolarHi10`, `B_Opt_TradeCap30`, and `B_Opt_TradeCap50`.

Regression characterization passed for discovery, filtering, ordering, comma splitting, whitespace, duplicate handling, unknown names, and empty filters.

## Evidence reuse decision

The failed environmental attempt is preserved unchanged at:

`C:\Users\luisfernando\Desktop\OSeMOSYS\OSTRAM_validation_evidence\refactor_preservation_87c3ce7_20260722_01`

Its manifest SHA-256 remains `227bbacfa102612e84cd05150740c299b83601932f126e827e36a0efb91dca04`.

The recent CLI evidence was not reused as the final preservation proof. Although its checkout and tree identity were verified and its artifact hashes matched the accepted report, its complete command and safety-monitor transcript could not be reconstructed. Static and safe-regression evidence from the stopped attempt was independently hash-verified, then rerun under this attempt. Fresh B1 and guarded B2 generation was performed in a new clean disposable checkout pinned to the candidate.

## Environment and commands

- OS: Windows `10.0.26200`
- Python: `3.10.20`, `C:\Users\luisfernando\anaconda3\envs\OSTRAM-env\python.exe`
- NumPy `2.2.6`, pandas `2.3.3`, PyYAML `6.0.3`, openpyxl `3.1.5`, otoole `1.1.5`

Canonical B1 generation:

```text
python -u -m ostram compile-inputs --scenarios A_Calibrated_BAU,A_Calibrated_BAU_Clipped,B_Optimised_VRE,B_Opt_Clipped,B_Opt_DirBidir,B_Opt_DirContractual,B_Opt_IndiaCosts,B_Opt_IndiaCostsFuel,B_Opt_SolarCapex130,B_Opt_SolarCapexHi,B_Opt_SolarCapexSpike,B_Opt_TradeCap15,B_Opt_TxCap150,C_Target_VRE,C_Target_VRE_Clipped
```

- Exit: 0
- Runtime: 1,362.393 seconds
- Scenarios: 15/15 completed
- Monitor: 579 polls, zero safety hits

Guarded B2 used the established `B2_Executing_OG_Model.py` orchestration with temporary dependency barriers and a strict subprocess allowlist. It permitted only 15 `otoole convert csv datafile` calls and 15 calls each to `preprocess_data.py`, `inject_DaysInDayType.py`, `patch_storage_delay.py`, `open_pwrbck_caps.py`, and `patch_reserve_margin_repair_careful_xlsx.py`.

- Allowlisted subprocesses: 90/90 returned 0
- Execution-stage barrier calls: 0
- Multiprocessing barrier calls: 0
- Runtime: 95.642 monitor seconds; 93.254 guarded-driver seconds
- Monitor: 41 polls, zero safety hits

The temporary monitor and JSON round-trip were tested with harmless no-op commands before generation. Disposable-tool failures involving PowerShell exit-code capture, UTF-8 redirection, the Windows platform probe, an overly narrow script-path allowlist, CSV-normalizer applicability, and the terminal `end;` parser marker were retained in evidence. Corrections were confined to temporary validation tooling; none changed repository behavior or accepted data.

The safe suite command was:

```text
python -m unittest discover -s tests\regression -p test_*.py -v
```

Static commands and their complete output are recorded in `10_static_gates_stdout.log` in the evidence directory.

## Static and test results

- Safe regression: 193 tests in 18.673 seconds; 190 passed, 3 skipped, 0 failed, 0 errors.
- Skips were optional only:
  - figure fixture: Matplotlib absent;
  - interconnection helpers: optional NumPy/Plotly visualization path;
  - visualization entrypoint smoke: optional visualization dependencies.
- Canonical versus historical CLI dispatch and first-boundary behavior: equivalent.
- Protected discovery: 20/20.
- Static cleanup acceptance: 16/16.
- Cleanup baseline comparison: 62 byte-exact, 2 normalized-exact.
- Protected tree: 2,356 files, 266,973,657 bytes, aggregate SHA-256 `a54cfc12e50c6c575ae823c9bf0fc381a24c207ab9041567fe97f24dafc24333`.
- Strict baseline self-comparison: no missing, extra, raw-drift, or normalized-drift files.
- AST: 123/123 tracked Python files parsed.
- Markdown: 45 files and 67 local link targets checked.
- Repository paths: 2,530 tracked paths checked.
- Deterministic EOL and import-safety tests: passed in the safe suite.
- `git diff --check`: passed.

No unexplained new skip occurred. Test safety characterization confirms no real solver or prohibited process boundary was invoked.

## Artifact equivalence

The repository comparator in `tests/regression/ostram_regression.py` supplied XLSX, CSV, and text normalization.

| Layer | Compared | Byte-exact | Normalized-exact | Notes |
| --- | ---: | ---: | ---: | --- |
| A1 completed workbooks | 60 | 0 | 60 | 60 metadata-only ZIP/core timestamp differences |
| Structure workbook | 1 | 0 | 1 | 1 metadata-only ZIP/core timestamp difference |
| `A2_Output_Params` | 660 | 660 | 660 | No missing or extra CSVs |
| `A2_Outputs_Params_otoole` | 960 | 960 | 960 | No missing or extra CSVs |
| B2 otoole datafiles | 15 | 15 | 15 | Exact |
| B2 concatenated input CSVs | 15 | 15 | 15 raw fallback | Structured CSV normalization is inapplicable because these intermediates contain legitimate duplicate CSV keys; raw equality is exact |
| B2 preprocessed TXTs | 15 | 15 | 15 | Exact |
| Storage-delay TXTs | 15 | 15 | 15 | Exact |
| Open-backstop TXTs | 15 | 15 | 15 | Exact |
| Final solver-consumed TXTs | 15 | 15 | 15 | Exact |
| Warning TXTs | 15 | 15 | 15 | Exact |

The seven selected B2 files per scenario total 105/105 byte-exact intermediates. Canonical final inventory was exactly 15 expected, 15 candidate, and 15 reference files, with zero missing and zero extra files.

All reference and candidate final sizes and hashes also match the committed accepted 15-file report.

## Semantic equivalence

Across the 15 final solver-consumed files, both candidate and reference contained:

- 810 parameter definitions;
- 1,437,655 parameter keys;
- 14,790 set definitions/blocks;
- 34,155 parsed set memberships.

Results:

- Parameter-definition drift: 0
- Parameter-key drift: 0
- Parameter-value drift: 0
- Set-definition drift: 0
- Set-membership drift: 0
- Duplicate parameter declarations: 0 candidate, 0 reference
- Duplicate parameter keys: 0 candidate, 0 reference
- Duplicate set declarations: 0 candidate, 0 reference
- Duplicate set memberships: 0 candidate, 0 reference

## Configuration and safety proof

Original `Config_MOMF_T1_AB.yaml` SHA-256:

`61e197cf175a63c1445c42fb9f71aea311d9a6db5d1642db9bf6a8feefe5bdbb`

Before guarded B2, PyYAML and the guard driver independently proved these were actual booleans with value `False`:

- `execute_model`
- `create_matrix`
- `concat_otoole_csv`
- `concat_scenarios_csv`

The temporary guarded config SHA-256 was `120ce53d260669d28a250163096b2aa076cf06280c56994908e6534a2502e022`. `storage_delay_model_output` was redirected to `.validation_runtime/osemosys_fast_preprocessed_storage_delay.txt` inside the disposable checkout. The original file was restored from an external byte backup and reverified at the original SHA-256. Configuration A was also restored automatically and verified at SHA-256 `64c6eb2d9dc55b91111cb969e63d2075ca9d267d2b1591ef3de23cab58e43261`; no A backup remained.

Across this rerun, 17 serialized monitor records covered 2,163.987 seconds and 930 polls with zero safety hits. Final audits found:

- Global solver processes: 0
- `.lp`, `.mps`, `.sav`, `.sol`, `.glp` files in the candidate: 0
- New or changed solver logs: 0
- Result or solution directories: 0
- New executable outputs: 0
- Matrix-path changes: 0
- Primary-repository generated artifacts: 0

## Evidence

Candidate checkout:

`C:\Users\luisfernando\AppData\Local\Temp\OSTRAM_refactor_preservation_87c3ce7_20260722_03`

Accepted reference checkout:

`C:\Users\luisfernando\Desktop\OSeMOSYS\OSTRAM_mainredo`

Accepted B1 evidence:

`C:\Users\luisfernando\Desktop\OSeMOSYS\OSTRAM_validation_evidence\b1_presolver_41a54e5`

Rerun evidence:

`C:\Users\luisfernando\Desktop\OSeMOSYS\OSTRAM_validation_evidence\refactor_preservation_87c3ce7_20260722_02`

- Source evidence files copied: 118
- Source evidence bytes: 2,881,109
- Source/destination SHA-256 checks: all matched
- Final evidence files, including manifests: 120
- `SHA256SUMS.txt` SHA-256: `d1c2ebe49567561d9098870cc9841627a5823aa1eea2ecd8f08072261a273360`
- `copy-verification.json` SHA-256: `5b796bce903815a6a66143945a0dc39c3c95c0f9b612806d19ecdda92a3676cb`
- Final comparison report SHA-256: `d31af4fa1f6f08dd1c14a2e794b86d337c3b9a4a717d8f4b9a6189d4ebeef46f`
- Final safety audit SHA-256: `8f472f3fc6c71bac486ff42e5a1575e64aaaedefa4caed459a5993189056a6b3`

The prior failed-evidence directory was neither overwritten nor relabeled. No disposable checkout was removed before evidence verification.

## Limitations and phase boundary

This is exact pre-solver preservation evidence. It does not claim CPLEX, GLPK, or other solver-output equivalence. The reference final files are generated, ignored files whose authority is established by their match to the committed hash report. XLSX raw bytes differ only in normalized archive/core metadata. The B2 concatenated input CSV layer is byte-exact, but its structured CSV normalizer is intentionally inapplicable because duplicate CSV keys are legitimate in that intermediate.

No production Python behavior, CLI entrypoint, scenario definition, classification, model input, workbook, interconnector datum, compiler selection, accepted output, or MOSOX path was changed. The repository is ready to proceed to `fix/interconnector-v18-source-of-truth`; phases 3–5 were not begun during this validation.
