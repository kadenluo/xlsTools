@echo off
path = %path%;.\tools;
@echo on
cd scripts && python xlsTools.py
pause
