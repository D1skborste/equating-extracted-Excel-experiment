@echo off

:: Temporary workaround! Point to the python.exe, even if you're not in a virtual env.
"C:\Users\PCMasterRace\AppData\Local\Python\bin\python.exe" "%~dp0compare_times.backup.py" %*

pause
