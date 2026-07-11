@echo off
REM Staged CPLEX solve — clipped baselines + cost/robustness sensitivities. Run from repo root.
REM One batch at a time; do NOT run concurrently with any other B1/B2 (see RUN_ORDER.md).
setlocal
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8
chcp 65001 >nul
set PY=C:\Users\luisfernando\anaconda3\envs\OSTRAM-env\python.exe
cd /d "%~dp0t1_confection"
"%PY%" -u B2_Executing_OG_Model.py --scenarios "A_Calibrated_BAU_Clipped,B_Opt_Clipped,C_Target_VRE_Clipped,B_Opt_SolarCapexHi,B_Opt_SolarCapex130,B_Opt_SolarCapexSpike,B_Opt_TradeCap15,B_Opt_TxCap150,B_Opt_IndiaCosts,B_Opt_IndiaCostsFuel"
endlocal
