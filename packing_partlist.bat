@echo off

call %userprofile%\scoop\apps\miniconda3\current\Scripts\activate.bat
REM call conda activate open-webui
call python packing_partlist.py

REM pause
