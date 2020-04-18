@echo off
path = %path%;.\tools;
@echo on
protoc --python_out=./ protocol/xls.proto
cd scripts && python xls2lua.py
pause