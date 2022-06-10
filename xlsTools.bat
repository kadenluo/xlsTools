@echo off
path = %path%;.\tools;
@echo on
cd scripts && python xls2lua.py
pause
