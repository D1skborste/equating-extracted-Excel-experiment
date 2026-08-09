@echo off

:: If problems arise, try pointing this to the python.exe, even if you're not in a virtual env.
set LAUNCH_METHOD=batch
"C:\Users\PCMasterRace\AppData\Local\Python\bin\python.exe" "%~dp0compare_times.py" %*

pause
