@echo off

REM path
set root=%USERPROFILE%\miniforge3
call %root%\Scripts\activate.bat %root%
call conda activate base

REM call conda activate open-webui
call python packing_partlist.py

REM pause
