@echo off

REM path
REM set root=%USERPROFILE%\miniforge3
REM call %root%\Scripts\activate.bat %root%
REM call conda activate base
REM call python sw_export.py

call uv run sw_export.py

pause
