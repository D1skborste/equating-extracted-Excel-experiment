@echo off

:: Temporary workaround! Point to the python.exe, even if you're not in a virtual env.
"C:\Users\MAKA6661\AppData\Local\Python\bin\python.exe" "%~dp0compare_times_excel.backup.py" %*

pause
