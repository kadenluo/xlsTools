@echo off
path = %path%;.\tools;
@echo on
protoc.exe --python_out=./ protocol/xls.proto
cd scripts && python xls2lua.py
pause