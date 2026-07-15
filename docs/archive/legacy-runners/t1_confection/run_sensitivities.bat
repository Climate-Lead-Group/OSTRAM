@echo off
REM ============================================================================
REM  run_sensitivities.bat  --  OSTRAM Phase-B CPLEX batch (CLG / OSTRAM)
REM ----------------------------------------------------------------------------
REM  Fire-and-forget: solves the 6 sensitivity scenarios SEQUENTIALLY (one CPLEX
REM  at a time), then concatenates, then runs the analysis. Wake up to
REM  sensitivity_report.txt + sensitivity_comparison.csv.
REM
REM  ROBUST: if a scenario fails/infeasible, it is logged and SKIPPED -- the
REM  batch CONTINUES to the next (it does NOT stop). concat + analysis then run
REM  on whatever solved.
REM
REM  Verified Config_MOMF_T1_AB.yaml state (do not need to change):
REM    parallel False | cplex_threads 4 | execute_model True | create_matrix True
REM    storage_delay_active True | strip_storage_active False | reuse_existing_sol False
REM
REM  Est. wall time: 6 solves x ~7-15 min (heat-soak dependent) ~= 55-90 min.
REM ============================================================================
chcp 65001
set PYTHONIOENCODING=utf-8
call conda activate OSTRAM-env
cd /d C:\Users\luisfernando\Desktop\OSeMOSYS\OSTRAM_clean\t1_confection
set FAILED=

echo ==================== [1/6] B_Opt_Clipped ====================
python B2_Executing_OG_Model.py --scenarios B_Opt_Clipped
if errorlevel 1 (echo [WARN] B_Opt_Clipped FAILED -- continuing & set FAILED=%FAILED% B_Opt_Clipped)

echo ==================== [2/6] B_Opt_TradeCap30 ====================
python B2_Executing_OG_Model.py --scenarios B_Opt_TradeCap30
if errorlevel 1 (echo [WARN] B_Opt_TradeCap30 FAILED -- continuing & set FAILED=%FAILED% B_Opt_TradeCap30)

echo ==================== [3/6] B_Opt_SolarCapexHi ====================
python B2_Executing_OG_Model.py --scenarios B_Opt_SolarCapexHi
if errorlevel 1 (echo [WARN] B_Opt_SolarCapexHi FAILED -- continuing & set FAILED=%FAILED% B_Opt_SolarCapexHi)

echo ==================== [4/6] B_Opt_TxCap150 ====================
python B2_Executing_OG_Model.py --scenarios B_Opt_TxCap150
if errorlevel 1 (echo [WARN] B_Opt_TxCap150 FAILED -- continuing & set FAILED=%FAILED% B_Opt_TxCap150)

echo ==================== [5/6] B_Opt_IndiaCosts ====================
python B2_Executing_OG_Model.py --scenarios B_Opt_IndiaCosts
if errorlevel 1 (echo [WARN] B_Opt_IndiaCosts FAILED -- continuing & set FAILED=%FAILED% B_Opt_IndiaCosts)

echo ==================== [6/6] B_Opt_IndiaCostsFuel ====================
python B2_Executing_OG_Model.py --scenarios B_Opt_IndiaCostsFuel
if errorlevel 1 (echo [WARN] B_Opt_IndiaCostsFuel FAILED -- continuing & set FAILED=%FAILED% B_Opt_IndiaCostsFuel)

echo ==================== concatenate all scenarios ====================
python concat_all_scenarios_2.py
if errorlevel 1 echo [WARN] concat FAILED

echo ==================== analysis ====================
python analyse_sensitivity.py
if errorlevel 1 echo [WARN] analysis FAILED (you can re-run: python analyse_sensitivity.py)

echo.
echo ==================== ALL DONE ====================
if defined FAILED (echo SCENARIOS THAT FAILED:%FAILED%) else (echo All 6 scenarios solved OK.)
echo Outputs: sensitivity_report.txt  and  sensitivity_comparison.csv
