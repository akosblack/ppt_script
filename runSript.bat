@echo off
setlocal

:: Get the current directory
set "curr_dir=%cd%"

:: Construct the new path for the robot file
set "new_path=%curr_dir%\script.py"

python %new_path%

endlocal
