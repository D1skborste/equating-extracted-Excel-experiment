@echo off
:: Explicitly points to the python.exe inside your .venv folder
set LAUNCH_METHOD=batch
"c:\Users\PCMasterRace\Documents\GitHub\venv\Scripts\python.exe" "%~dp0compare_times.py" %*

pause