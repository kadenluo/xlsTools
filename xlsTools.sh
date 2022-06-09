
./tools/protoc --python_out=./ ./protocol/xls.proto
cd scripts && python3 xls2lua.py
