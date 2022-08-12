@echo off
set /p USERNAME="AO3 Username: "
set /p PASSWORD="AO3 Password: "
set /p BROWSER="Browser (chrome, firefox): "
set /p BKMK="Include bookmark status? (y/n): "

python saveHistory.py %USERNAME% %PASSWORD% %BROWSER% %BKMK%
pause