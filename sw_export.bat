@echo off

call %userprofile%\scoop\apps\miniconda3\current\Scripts\activate.bat
call python sw_export.py

pause
