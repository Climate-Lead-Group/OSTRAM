# Release-readiness evidence

This record supplements the solver-free integration evidence in `provenance.md`. Generated
model, solver, CSV, snapshot, and dashboard files remain ignored under
`workspace/release-readiness/profiles/unescap/`; none is committed.

## Windows environment and one authorized CBC run

- CBC: 2.10.13, build date 2026-03-12.
- Conda: 24.5.0; environment `OSTRAM-env`.
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

The OSTRAM parent process returned 0, but CBC did **not** find a feasible solution:

```text
PrimalInfeasible objective 559384.1792 - 137927 iterations
Result - Linear relaxation infeasible
```

The solution header is:

```text
Infeasible - objective value 559384.17921019
```

Otoole independently warned that the CBC solution contains decision variables out of
bounds. The exact boundary is the active execution chain for
`B_Optimised_VRE`: the compiled domain validates, preprocessing and the storage-delay,
open-backstop, and careful reserve-margin patches complete, GLPK writes the 90,384-column
LP, and CBC finds its linear relaxation infeasible. The careful reserve-margin patch also
records 10 warnings, including zero early investment caps for `PWRNGSBGDXX` and
`PWRNGSINDEA`; those warnings are evidence, not proof of the infeasibility's modelling
cause. No assumption was changed and no second solver run was made.

The generated evidence is reproducibly identified by:

| Ignored artifact | SHA-256 |
|---|---|
| `execution/Executables/B_Optimised_VRE_0/Pre_processed_B_Optimised_VRE_0_StorageDelayN5_OpenBCK_RMCarefulXLSX_output.sol` | `78e0558f993f66cfb1cf6b0c56cc73b87b67da6fb428f64f964d23d990698899` |
| `execution/OSTRAM_StorageDelay_Combined_Inputs_Outputs.csv` | `5ee2d9a4a134bcfc006a083319fc8cf9684225d94837b09249d9a4c127dd93b3` |
| `reports/unescap.html` | `dbb1d20c9817d5fc2e1725b3149e3e7d934e697839bae2796f84d50fffba66da` |

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
  prepatch inputs. All 15 regenerated final `.txt` files were byte-exact to the accepted
  artifacts, including their explicit CRLF convention; no matrix or solver ran.
- Compileall and imports passed from outside the repository caller CWD.
- Canonical help and `inspect-resources` routes passed from outside the repository CWD.
- The real solver-boundary domain revalidated at 90 technologies and 49 fuels, with only
  `PWRSHPINDEA` and the six declared ELC dispatch fuels added to the 89/43 seed contract.
- The macOS workflow parses as YAML, all three embedded Python blocks compile, and its
  compile-only domain block ran locally against the real prepared output. The actual
  `macos-latest` ARM64 job remains necessarily unexecuted until the branch is pushed.
