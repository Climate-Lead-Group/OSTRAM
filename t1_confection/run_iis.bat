@echo off
REM ============================================================================
REM run_iis.bat
REM ----------------------------------------------------------------------------
REM Runs CPLEX conflict refiner (IIS) on the existing infeasible LP from
REM the BAU_0 NoStorage run, using the same cplex.exe that B2 uses.
REM
REM USAGE
REM   1. Place this .bat alongside iis_cmds.txt in any scratch folder.
REM   2. Open cmd in that folder and run:   run_iis.bat
REM      (or just double-click)
REM
REM FILES WRITTEN (all NEW, nothing existing is overwritten):
REM   - iis_run.log     (in the folder you run from)
REM   - Pre_processed_BAU_0_NoStorage_output_conflict.clp
REM     (written next to the original .lp file)
REM
REM FILES UNTOUCHED:
REM   - Pre_processed_BAU_0_NoStorage_output.lp
REM   - Pre_processed_BAU_0_NoStorage_output.sol
REM   - Pre_processed_BAU_0_NoStorage_output.feasopt.sol
REM   - Pre_processed_BAU_0_NoStorage_output.cplex.log
REM
REM EXPECTED RUNTIME: 15-45 minutes (4 min optimize + ~10-40 min conflict)
REM
REM IF "conflict" THROWS A SYNTAX ERROR ON CPLEX 22.1:
REM   Open iis_cmds.txt and replace the line   conflict
REM   with                                     tools conflict
REM ============================================================================

cplex < iis_cmds.txt

echo.
echo --- Done ---
echo Conflict file:
echo   C:\Users\luisfernando\Desktop\OSeMOSYS\OSTRAM_storage_debug_try2\t1_confection\Executables\BAU_0\Pre_processed_BAU_0_NoStorage_OpenBCK_output_conflict.clp
echo Run log:
echo   %CD%\iis_run.log
