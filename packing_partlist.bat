@echo off

REM path
REM set root=%USERPROFILE%\miniforge3
REM call %root%\Scripts\activate.bat %root%
REM call conda activate base

REM call conda activate open-webui
REM call python packing_partlist.py

call uv run packing_partlist.py

REM pause
