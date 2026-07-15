@echo off
REM Staged CPLEX solve — baselines. Run from repo root. One batch at a time (see RUN_ORDER.md).
setlocal
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8
chcp 65001 >nul
set PY=C:\Users\luisfernando\anaconda3\envs\OSTRAM-env\python.exe
cd /d "%~dp0t1_confection"
"%PY%" -u B2_Executing_OG_Model.py --scenarios "A_Calibrated_BAU,C_Target_VRE,B_Optimised_VRE"
endlocal
