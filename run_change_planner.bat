@echo off
setlocal
cd /d "%~dp0"

echo WU Assumptions Generator - Report Change Planner
echo ------------------------------------------------
echo This utility creates an implementation plan. It does NOT modify the app.
echo.

if exist ".venv\Scripts\python.exe" goto usevenv
where py >nul 2>nul
if not errorlevel 1 goto usepy
where python >nul 2>nul
if not errorlevel 1 goto usepython

echo ERROR: Python was not found.
echo Run run_app.bat once, or install Python 3.11 or newer.
goto done

:usevenv
".venv\Scripts\python.exe" report_change_planner.py --project-root "%CD%"
goto done

:usepy
py -3 report_change_planner.py --project-root "%CD%"
goto done

:usepython
python report_change_planner.py --project-root "%CD%"

goto done

:done
echo.
pause
endlocal
