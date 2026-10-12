# UNESCAP authoring and acceptance guide

This guide covers authoring and testing the Windows training exercises. Run
commands from an Anaconda or Miniforge prompt using `cmd.exe`. The commands resolve
project resources independently of the caller's current directory, but keeping the
repository root as the prompt location makes Git review easier.

## Install a clean checkout on Windows

```text
git clone https://github.com/Climate-Lead-Group/OSTRAM.git OSTRAM
cd OSTRAM
```

If the repository already exists, use `git fetch origin`, `git switch main` and
`git pull`. Confirm the branch with `git branch --show-current`.

Install the supported environment once:

```text
conda env create -n ostram-unescap -f environment.yaml
conda activate ostram-unescap
set PYTHONUTF8=1
set PYTHONHASHSEED=0
python -m pip install -e . --no-deps
python -m pip check
python -m ostram --help
cbc -stop
glpsol --version
```

Repeat `conda activate ostram-unescap` and the two `set` commands in each new
prompt. These match the encoding and deterministic hash seed used for acceptance.

If this dedicated `ostram-unescap` environment already exists, replace the create command with:

```text
conda env update -n ostram-unescap -f environment.yaml --prune
```

`cbc -stop` must print the CBC welcome banner and version, then exit. Report the full
output if Windows cannot find `cbc`; do not substitute another solver.

## Establish a clean Windows baseline

Use a dedicated ignored workspace so training state never mixes with another profile:

```text
python -m ostram --workspace workspace/natalia example prepare unescap
python -m ostram --workspace workspace/natalia --profile unescap inspect-resources
python -m ostram --workspace workspace/natalia --profile unescap run --env-name ostram-unescap --scenarios "A_Calibrated_BAU,B_Optimised_VRE" --compile-only --skip-pull
```

Preparation is guarded. Repeating the first command refuses to replace existing state;
the explicit reset is:

```text
python -m ostram --workspace workspace/natalia example prepare unescap --reset
```

The baseline interconnector authority is exactly **2.496 GW** in 2023-2028. The active
B scenario intentionally relaxes `TotalAnnualMaxCapacity` to **9999.0**. A 2.5 GW value
is an exercise edit or prose approximation, not the committed seed.

Run A and B through CBC, export and cost reconciliation, then prepare and run C
using the completed A result in this same workspace:

```text
python -m ostram --workspace workspace/natalia --profile unescap run --env-name ostram-unescap --scenarios "A_Calibrated_BAU,B_Optimised_VRE" --skip-pull
python -m ostram --workspace workspace/natalia --profile unescap run --env-name ostram-unescap --scenarios "C_Target_VRE" --skip-pull
python -m ostram --workspace workspace/natalia example report unescap --capture windows-cbc-baseline
python -m ostram --workspace workspace/natalia example report unescap
```

Select the decision scenarios explicitly. A materializes support BAU's restrictions
without compiling or solving BAU; C requires a completed A result. Do not supply an
accepted full-model result as a training input. Capture a report only after the run
returns success with an optimal CBC solution and successful cost reconciliation.
The profile enables backstop capacity; inspect its generation and capacity in the
exports when interpreting feasibility.

The final command prints the real dashboard path. All solver results, captures, logs,
and HTML dashboards remain below `workspace/natalia/profiles/unescap/` and are ignored.

## Walk Exercise A and Exercise B

Open `exercises/training.html` first, then walk:

1. `exercises/add-country.html` (Exercise A). Start with its guarded `--reset`, generate
   and merge the MMR country template, run the documented A1+A2-only command, use
   `scenario sync-country --country MMR --dry-run`, repeat without `--dry-run`, and
   validate MMR. Do not reset after the country merge.
2. `exercises/add-interconnector.html` (Exercise B). Continue from Exercise A. Use
   `inspect-resources` to locate and edit the prepared workspace workbook, while the
   taxonomy, extension decision, and scenario YAML teaching edits remain the explicit
   profile files named in the exercise.
3. Use `--compile-only` while checking workbook/config edits. Invoke CBC only when the
   exercise specifically reaches its solve step.
4. Capture each solved direction with a descriptive label. When `forward`, `reverse`,
   and `bidirectional` exist, run:

```text
python -m ostram --workspace workspace/natalia example report unescap --compare forward,reverse,bidirectional
```

Every command's authoritative flags are visible through `--help`, for example:

```text
python -m ostram --workspace workspace/natalia example report unescap --help
python -m ostram --workspace workspace/natalia --profile unescap country template --help
python -m ostram --workspace workspace/natalia --profile unescap scenario sync-country --help
```

## Edit and validate the exercise HTML

Training authors may edit the committed authoring assets under `examples/unescap/exercises/`,
this guide, the operations table, and intentional profile seed/config authorities needed
by the exercise. After editing HTML, run:

```text
python -m unittest tests.regression.test_release_readiness.TrainingAssetTests -v
python -m unittest discover -s tests -p "test_*.py"
git diff --check
git status --short
```

The focused test validates local HTML links/fragments and rejects retired or invented
commands. In a browser, click every navigation link and copy button as an additional
authoring check.

Committed profile seeds live under `examples/unescap/` and are reviewable source.
Mutable preparation, compiled models, logs, `.sol` files, combined result CSVs, report
captures, and dashboards live under `workspace/` and must not be committed. Before a
commit, `git status --short --ignored` should show generated state only as ignored.
Never add a workaround Python script to an exercise directory. Record the exact command,
run log path, traceback, and expected behavior, then report it as an engine defect.

## macOS acceptance in GitHub Actions

macOS acceptance comes from the pull request's GitHub Actions run. It needs no physical
Mac, and the workflow YAML is not edited for it.

1. Open the integration pull request and select **Checks**, or open the repository's
   **Actions** tab and choose **UNESCAP macOS ARM64 acceptance**.
2. The required job is **macOS ARM64 UNESCAP acceptance**. Open the job to inspect each
   step and its logs.
3. Confirm the environment step reports a macOS version, `uname -m` as `arm64`, the
   expected Python, and CBC's version banner.
4. Confirm the complete solver-free suite, preparation, B compile-only, compiled-domain
   validation, reduced CBC smoke, report capture, and dashboard validation steps pass.
5. In the run's **Artifacts** section, download `unescap-macos-arm64-acceptance`. Open
   `acceptance-summary.md` and `acceptance-summary.json`, then open the included report
   HTML locally. Check BGDXX, INDEA, and TRNBGDXXINDEA coverage.
6. If a check fails, record the run URL, job name, failing step, and first relevant error,
   and report those details as a defect. Do not create a local legacy script.

The automated gate does not claim interactive Microsoft Excel-on-Mac validation. That
remains outside this acceptance boundary.

## Acceptance report template

```text
UNESCAP training acceptance
Branch / commit:

Windows
- Python:
- CBC version:
- Solver-free tests:
- Prepare / reset:
- B compile-only:
- B CBC status and duration:
- Capture label:
- Dashboard path and BGDXX / INDEA / TRN coverage:
- Exercise A walk:
- Exercise B walk:

macOS GitHub Actions
- Run URL:
- Job: macOS ARM64 UNESCAP acceptance
- uname -m (arm64):
- Python / CBC:
- Solver-free tests:
- Prepare / compile / domain:
- CBC status:
- Artifact downloaded:
- Dashboard opened and coverage checked:

Interactive Excel-on-Mac: not part of automated acceptance
Problems (run URL, step, log excerpt):
```
