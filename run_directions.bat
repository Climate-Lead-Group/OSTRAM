@echo off
REM Staged CPLEX solve — interconnector direction scenarios. Run from repo root, LAST.
REM One batch at a time; do NOT run concurrently with any other B1/B2 (see RUN_ORDER.md).
setlocal
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8
chcp 65001 >nul
set PY=C:\Users\luisfernando\anaconda3\envs\OSTRAM-env\python.exe
cd /d "%~dp0t1_confection"
"%PY%" -u B2_Executing_OG_Model.py --scenarios "B_Opt_DirBidir,B_Opt_DirContractual"
endlocal
