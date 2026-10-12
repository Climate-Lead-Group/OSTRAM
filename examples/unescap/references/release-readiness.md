# Release-readiness evidence

## 2026-09-22 isolated training reproduction repair

This record supersedes the historical training solver outcome below. The repair
starts at commit `6ff8fc91c89b81a5e44cbf6798f0ec30a0fbc77a` and is merged into
`main`. The clean-clone commands are in
[the authoring and acceptance guide](../AUTHORING_AND_ACCEPTANCE.md).

The original maintained inputs reproduce the missing Bangladesh gas emissions
error and all three extension plus sixteen manual-fix failures in support BAU,
A and B. The confirmed causes, activity basis and bounded changes are recorded in
[`provenance.md`](provenance.md#bangladesh-gas-emissions-and-training-a3-checks).
The repaired A3 stages pass 53 extension and 49 manual-fix checks per prepared
scenario; all 24 compared workbook cell-value sets are unchanged. Negative cases
still reject missing small hydro and missing/duplicate Bangladesh gas targets.

Final A and B completed native compilation, CBC solving, export and cost
reconciliation from an empty workspace. C then completed the same chain using
that workspace's freshly generated A results. All three CBC headers declare
optimal solutions. Windows Python 3.10.21, CBC 2.10.13, GLPK 5.0 and otoole 1.1.5
were used; `pip check` passes. Training uses CBC primal tolerance `1e-10` for A/B
and `1e-12` for C, dual tolerance `1e-8`, and both random seeds 12345.

Maximum new/installed annual backstop capacity (GW):
A_Calibrated_BAU: 0/0; B_Optimised_VRE: 0/0; C_Target_VRE: 0/0.
Maximum saved bound violation:
A_Calibrated_BAU: 0; B_Optimised_VRE: 0; C_Target_VRE: 7.42e-11.

The reconciled total includes GLPK's omitted objective constant; it is compared
with that constant plus the CBC objective. Every cost difference is within
`1e-7` relative tolerance. Raw exports are retained. The independent LP check
evaluates every row and bound at `1e-7` scaled tolerance because CBC's saved
text has about eight significant digits. Material CBC `**` flags are rejected;
negative postsolve round-off within `1e-8` of zero is allowed. Otoole 1.1.5 emits
its generic infeasible-solution warning for C because it sees those `**` markers;
the retained solution and independent residual check establish their magnitude
above. The warning is preserved in the log. All three final cases have zero
backstop generation and capacity; the profile's backstop remains enabled.

Each compiled domain remains GLOBAL, 90 technologies, 49 fuels, 28 years
(2023–2050), 20 timeslices and four storage technologies, with only BGDXX/INDEA
country content and physical interconnector `TRNBGDXXINDEA`. Both URN technologies
remain present and no NUC code is introduced. All 56 Bangladesh gas emission
coefficients match the compiled activity ratios. The existing dashboard/capture
acceptance checks pass for all three fresh scenarios and both country regions.

The full-profile comparison freshly regenerates **all 17** accepted final text
inputs and LPs: every SHA-256 matches the specified reproduction evidence.
This covers 1,644,794 compiled coefficient comparisons by byte identity.
Full-profile inputs, configuration, maintained
model and solver commands/settings are unchanged. A fresh full-profile A solve
supplied the C compilation dependency; the other accepted solves are reused
only as comparison evidence for unchanged LPs/settings. Accepted artifacts were
never runtime inputs. Production, delivery, accepted outputs and PR #35 were
not edited.

The final solver-free suite runs 317 tests with exactly two errors, both the
known stale 15-versus-17 contract:
`test_governed_manifest_binds_root_and_derived_outputs` and
`test_narrow_ignore_rules_and_scope`. There are no other failures/errors.
This branch deliberately leaves that correction to
PR #35; integrating its test-contract correction is still required for a green
combined suite. The training repair has no runtime dependency on the BAU fix.

Detailed commands, logs, hashes, LP residuals, per-technology/year backstop use,
cost reconciliation and the 17-case comparison were retained outside the
repository in the repair records: `repair-summary.json`,
`training-acceptance-*.json`, `training-report-acceptance.json`,
`full-invariance-results.json`, `a3-before-details.json`,
`a3-after-details.json` and `final-regressions.log`, together with the
generated training outputs and the full comparison artifacts. None of these
is committed. This is Windows acceptance; no macOS run is claimed.

## Historical release-readiness record

This record supplements the solver-free integration evidence in `provenance.md`. Generated
model, solver, CSV, snapshot, and dashboard files remain ignored under
`workspace/release-readiness/profiles/unescap/`; none is committed.

## Windows environment and one authorized CBC run

- CBC: 2.10.13, build date 2026-03-12.
- Conda: 24.5.0; the environment created from `environment.yaml`.
- Scenario: `B_Optimised_VRE` only; no full-model solver was run.
- Canonical process return code: 0.
- Total pipeline duration: 618.758 seconds (10m 18s).
- Stage durations: A1 40s, A2 7s, A3 49s, B1 18s, B2 8m 08s.
- CBC duration reported by CBC: 437.35 wall-clock seconds.
- CBC iterations: 137,927.

The first invocation returned 1 in 0.202 seconds because this PowerShell process did not
have `conda` on `PATH`. It stopped before A1, matrix creation, or CBC. After activating
the existing supported Conda paths, the single authorized solver process ran.

## Genuine solver boundary

The OSTRAM parent process returned 0. CBC found no feasible solution: it reported its
linear relaxation infeasible after 137,927 iterations, and the solution header reads
`Infeasible`.

Otoole independently warned that the CBC solution contains decision variables out of
bounds. The exact boundary is the active execution chain for
`B_Optimised_VRE`: the compiled domain validates, preprocessing and the storage-delay,
open-backstop, and careful reserve-margin adjustments complete, GLPK writes the 90,384-column
LP, and CBC finds its linear relaxation infeasible. The careful reserve-margin adjustment also
records 10 warnings, including zero early investment caps for `PWRNGSBGDXX` and
`PWRNGSINDEA`; those warnings are evidence, not proof of the infeasibility's modelling
cause. No assumption was changed and no second solver run was made.

The runtime now validates CBC's solution header before otoole conversion and reporting;
zero-exit infeasible results fail closed. The macOS acceptance workflow applies the same
check. This is a solver-status integration correction, not a model-authority change.

## Report route evidence

After repointing the editable environment install to this checkout, the canonical capture
and report commands succeeded from a caller directory outside the repository root using
the descriptive label `windows-cbc-release-baseline`. The generated HTML has the expected
doctype, `report` and `ostram-profile-data` elements, valid internal links, BGDXX and INDEA
metadata, and TRNBGDXXINDEA series for 2023-2050. Because the source solution is
infeasible, that HTML is route/structure evidence only and is **not** accepted model-result
evidence.

Actual M1 execution remains pending until the branch is pushed and the pull-request
workflow starts. Interactive Microsoft Excel-on-Mac is outside the automated boundary.

## Final static and solver-free gates

- Complete solver-free suite: 252 tests passed in 116.430 seconds.
- Focused portability, report, capture, B2, and governed-newline tests: passed.
- The changed reserve-margin writer was replayed over the canonical 15 full-model
  pre-adjustment inputs. All 15 regenerated final `.txt` files were byte-exact to the accepted
  artifacts, including their explicit CRLF convention; no matrix or solver ran.
- Compileall and imports passed from outside the repository caller CWD.
- Canonical help and `inspect-resources` routes passed from outside the repository CWD.
- The real solver-boundary domain revalidated at 90 technologies and 49 fuels, with only
  `PWRSHPINDEA` and the six declared ELC dispatch fuels added to the 89/43 seed contract.
- The macOS workflow parses as YAML, all three embedded Python blocks compile, and its
  compile-only domain block ran locally against the real prepared output. The actual
  `macos-latest` ARM64 job remains necessarily unexecuted until the branch is pushed.
