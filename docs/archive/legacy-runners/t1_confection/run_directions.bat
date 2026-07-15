@echo off
REM ============================================================================
REM  run_directions.bat  --  OSTRAM interconnector-direction batch (CLG / OSTRAM)
REM ----------------------------------------------------------------------------
REM  Solves the 2 interconnector-direction scenarios SEQUENTIALLY (one CPLEX at a
REM  time), then concatenates ALL scenarios, then runs the analysis.
REM
REM    [1/2] B_Opt_DirBidir       -- validation anchor; must reproduce B_Opt_Clipped
REM                                  (2,115,082 M USD). All corridors bidirectional.
REM    [2/2] B_Opt_DirContractual -- interconnectors locked to real contractual
REM                                  directions (see reference/
REM                                  interconnector_direction_references.md).
REM
REM  ROBUST: if a scenario fails/infeasible it is logged and SKIPPED; the batch
REM  CONTINUES. concat + analysis then run on whatever solved.
REM
REM  Prereq (already verified by Claude via glpsol --check before handoff):
REM    Config_MOMF_T1_AB.yaml -> execute_model True | create_matrix True |
REM    cplex_threads 4 | storage_delay_active True | reuse_existing_sol False
REM
REM  Est. wall time: 2 solves x ~7-12 min ~= 15-25 min (both are loose, ceiling-
REM  only LPs; DirContractual only removes flow options, so no harder than B_Opt).
REM ============================================================================
chcp 65001
set PYTHONIOENCODING=utf-8
call conda activate OSTRAM-env
cd /d C:\Users\luisfernando\Desktop\OSeMOSYS\OSTRAM_clean\t1_confection
set FAILED=

echo ==================== [1/2] B_Opt_DirBidir ====================
python B2_Executing_OG_Model.py --scenarios B_Opt_DirBidir
if errorlevel 1 (echo [WARN] B_Opt_DirBidir FAILED -- continuing & set FAILED=%FAILED% B_Opt_DirBidir)

echo ==================== [2/2] B_Opt_DirContractual ====================
python B2_Executing_OG_Model.py --scenarios B_Opt_DirContractual
if errorlevel 1 (echo [WARN] B_Opt_DirContractual FAILED -- continuing & set FAILED=%FAILED% B_Opt_DirContractual)

echo ==================== concatenate all scenarios ====================
python concat_all_scenarios_2.py
if errorlevel 1 echo [WARN] concat FAILED

echo ==================== analysis ====================
python analyse_sensitivity.py
if errorlevel 1 echo [WARN] analysis FAILED (you can re-run: python analyse_sensitivity.py)

echo.
echo ==================== ALL DONE ====================
if defined FAILED (echo SCENARIOS THAT FAILED:%FAILED%) else (echo Both direction scenarios solved OK.)
echo Validation: B_Opt_DirBidir should equal B_Opt_Clipped (2,115,082 M USD).
echo Outputs: sensitivity_report.txt  and  sensitivity_comparison.csv
