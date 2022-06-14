#!/usr/bin/python3
# -*-  coding:utf-8 -*-

import os
import re
import sys
import time
import datetime
import xlrd
import json
import argparse
import traceback
from luaparser import ast
from operator import itemgetter

class Logger():
    def __init__(self):
        pass

    def debug(self, msg):
        print("[DEBUG] {}".format(msg))

    def info(self, msg):
        print("[INFO] {}".format(msg))

    def error(self, msg):
        print("[ERROR] {}".format(msg))

datemode = 0 # 时间戳模式 0: 1900-based, 1: 1904-based
BOOL_YES = ["yes", "1", "是"]
BOOL_NO = ["", "nil", "0", "false", "no", "none", "否", "无"]

class Converter:
    _config = {} # 输入配置
    _indent = "    " #缩进
    _cachefile = "./.cache"
    _logger = None
    def __init__(self, config, logger):
        self._logger = logger
        assert(config.type == "all" or config.type == "lua" or config.type == "json")
        assert(config.export == "server" or config.export == "client")
        self._config = config

    def _toLua(self, data, level=1):
        lines = []
        if isinstance(data, list):
            items = []
            for value in data:
                value = self._toLua(value, level+1)
                items.append("%s%s"%(self._indent*level, value))
            lines.append("{\n%s\n%s}"%(", \n".join(items), self._indent*(level-1)))
        elif isinstance(data, dict):
            items = []
            for (key, value) in data.items():
                if isinstance(key, int):
                    key = "[%d]"%(key)
                elif isinstance(key, str):
                    pass
                else:
                    raise Exception("Error: {}({}) can't be a key.".format(key, type(key)))
                value = self._toLua(value, level+1)
                items.append("{}{} = {}".format(self._indent*level, key, value))
            lines.append("{\n%s\n%s}"%(", \n".join(items), self._indent*(level-1)))
        elif isinstance(data, str):
            return '[[%s]]'%(data)
        elif isinstance(data, bool):
            return 'true' if data else 'false'
        else:
            return str(data)
        return ", \n".join(lines)

    def save(self, output_type, filename, data):
        out_dir = os.path.join(self._config.output_dir, output_type)
        if not os.path.exists(out_dir):
            os.makedirs(out_dir, mode=0o755)
        filepath = os.path.join(out_dir, "{}.{}".format(filename, output_type))
        with open(filepath, 'wb') as f:
            f.write(data.encode('utf-8'))

    def convertAll(self):
        try:
            history = {}
            if not self._config.force:
                if os.path.exists(self._cachefile):
                    with open(self._cachefile) as f:
                        history = json.load(f)

            allfiles = {}
            for filename in os.listdir(self._config.input_dir):
                if filename.startswith("~"):
                    continue

                filepath = os.path.join(self._config.input_dir, filename)
                mtime = os.stat(filepath).st_mtime
                allfiles[filename] = mtime

                if (filename not in history) or (history[filename] != mtime):
                    self.convertFile(filename)
                    history[filename] = mtime
                    with open(self._cachefile, "w") as f:
                        json.dump(history, f, indent=4)

            # 清理cache
            delkeys = []
            for (filename, mtime) in history.items():
                if filename not in allfiles:
                    delkeys.append(filename)
            for k in delkeys:
                del history[filename]
            with open(self._cachefile, "w") as f:
                json.dump(history, f, indent=4)

            self._logger.info("success!!!")
        except Exception as ex:
            self._logger.error(traceback.format_exc())

    def convertFile(self, filename):
        filepath = os.path.join(self._config.input_dir, filename)
        wb = xlrd.open_workbook(filepath)
        for sheet in wb.sheets():
            self._logger.info("convert {}({})...".format(filename, sheet.name))
            self._convertSheet(sheet)

    def _convertSheet(self, sheet):
        nrows = sheet.nrows
        ncols = sheet.ncols
        classname = sheet.name
        assert ((nrows > 3) and (ncols > 1))

        mainkey = None
        field2index = {}
        uniqueFields = {} # 用于校验field重复问题
        for col in range(ncols):
            desc = self._getCellString(sheet.cell(0, col)).strip(' ')
            name = self._getCellString(sheet.cell(1, col)).strip(' ')
            vtype = self._getCellString(sheet.cell(2, col)).strip(' ')
            export = self._getCellString(sheet.cell(3, col)).strip(' ').lower()

            if len(name) == 0:
                continue

            assert(export == "all" or export == "" or export == "client" or export == "server")
            if export == "all" or export == "" or self._config.export == export:
                pass
            else:
                continue

            if name.startswith('*'):
                name = name.strip('*')
                assert(mainkey == None)
                mainkey = name
            assert(name not in uniqueFields)
            uniqueFields[name] = True
            field2index[col] = {"desc":desc, "name":name, "type":vtype, "levels":name.split('#')}

        uniqueFields.clear()
        if mainkey is None:
            result = []
        else:
            result = {}
        for row in range(4, nrows):
            item = {"_meta":{"isdict":True}}
            fields = {}
            for col in range(ncols):
                if col not in field2index:
                    continue
                meta = field2index[col]
                value = self._getCellValue(sheet.cell(row, col), meta["type"]) 
                key = meta["name"]
                if len(meta["levels"]) == 1:
                    item[key] = value
                else:
                    fields[key] = value
            self._convertRow(item, fields)
            item = self._fixLevelType(item)
            if isinstance(result, list):
                result.append(item)
            else:
                k = item[mainkey]
                del item[mainkey]
                result[k] = item

        #self._logger.info(json.dumps(result, indent=4))

        if self._config.type == "all" or self._config.type == "lua":
            luaTable = self._toLua(result)
            code = "local data = %s\n\nreturn data" % (luaTable)
            ast.parse(code)
            self.save("lua", classname, code)

        if self._config.type == "all" or self._config.type == "json":
            code = json.dumps(result, indent=4)
            self.save("json", classname, code)

    def _convertRow(self, result, fields):
        if len(fields) == 0 :
            return

        keys = sorted(fields.keys())
        total = len(keys)

        idx = 0
        childfields = {}
        while idx < total:
            key = keys[idx]
            value = fields[key]

            path = key.split('#')
            childitem = {"_meta":{"isdict":not path[-1].isdigit()}}
            prefix = "#".join(path[:-1])
            while idx < total:
                k = keys[idx]
                if k.startswith(prefix):
                    p = k.split('#')
                    childitem[p[-1]] = fields[k]
                else:
                    idx = idx - 1
                    break
                idx = idx + 1

            if len(path) > 2:
                childfields[prefix] = childitem
            else:
                result[prefix] = childitem 
            idx = idx + 1

        self._convertRow(result, childfields)

    def _fixLevelType(self, item):
        meta = item["_meta"]
        del item["_meta"]

        for (key, value) in item.items():
            if isinstance(value, dict):
                item[key] = self._fixLevelType(value)

        if meta["isdict"]:
            pass
        else:
            tmp = []
            for k in sorted(item.keys()):
                tmp.append(item[k])
            item = tmp
        return item

    def _getCellValue(self, cell, vtype):
        if vtype == "int":
            return self._getCellInt(cell)
        elif vtype == "float":
            return self._getCellFloat(cell)
        elif vtype == "bool":
            return self._getCellBool(cell)
        elif vtype == "string": # string
            return self._getCellString(cell)
        elif vtype == "list(int)":
            return self._getCellListForInt(cell)
        elif vtype == "list(float)":
            return self._getCellListForFloat(cell)
        elif vtype == "list(bool)":
            return self._getCellListForBool(cell)
        elif vtype == "list(string)":
            return self._getCellListForString(cell)
        else:
            raise Exception("This type is invalid. %s" % vtype)

    def _getCellListForInt(self, cell):
        if cell.ctype == xlrd.XL_CELL_TEXT:
            result = []
            value = cell.value.strip(' ').lstrip('[').rstrip(']')
            if len(value) == 0:
                return []
            for v in value.split(','):
                result.append(int(v))
            return result
        else:
            raise Exception("Error: invalid cell type. type:{}".format(cell.ctype))

    def _getCellListForFloat(self, cell):
        if cell.ctype == xlrd.XL_CELL_TEXT:
            result = []
            value = cell.value.strip(' ').lstrip('[').rstrip(']')
            if len(value) == 0:
                return []
            for v in value.split(','):
                result.append(float(v))
            return result
        else:
            raise Exception("Error: invalid cell type. type:{}".format(cell.ctype))

    def _getCellListForBool(self, cell):
        if cell.ctype == xlrd.XL_CELL_TEXT:
            result = []
            value = cell.value.strip(' ').lstrip('[').rstrip(']')
            if len(value) == 0:
                return []
            for v in value.split(','):
                v = v.lower()
                if v in BOOL_NO:
                    v = False
                elif value in BOOL_YES: 
                    v =  True
                else:
                    raise Exception("Error: can't switch the value to bool. value:{}".format(v))
                result.append(v)
            return result
        else:
            raise Exception("Error: invalid cell type. type:{}".format(cell.ctype))

    def _getCellListForString(self, cell):
        if cell.ctype == xlrd.XL_CELL_TEXT:
            value = cell.value.strip(' ').lstrip('[').rstrip(']')
            if len(value) == 0:
                return []
            return value.split(',')
        else:
            raise Exception("Error: invalid cell type. type:{}".format(cell.ctype))
        
    def _getCellString(self, cell):
        if cell.ctype == xlrd.XL_CELL_EMPTY:
            return ""
        elif cell.ctype == xlrd.XL_CELL_TEXT:
            return cell.value
        elif cell.ctype == xlrd.XL_CELL_NUMBER:
            return str(cell.value)
        elif cell.ctype == xlrd.XL_CELL_DATE:
            dt = xlrd.xldate.xldate_as_datetime(cell.value, datemode)
            return "%s" % dt
        elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
            return "true" if cell.value else "false"
        else:
            raise Exception("Error: invalid cell type. type:{}".format(cell.ctype))
    
    def _getCellInt(self, cell):
        if cell.ctype == xlrd.XL_CELL_EMPTY:
            return int(0)
        elif cell.ctype == xlrd.XL_CELL_TEXT or cell.ctype == xlrd.XL_CELL_NUMBER:
            return int(cell.value)
        elif cell.ctype == xlrd.XL_CELL_DATE:
            dt = xlrd.xldate.xldate_as_datetime(cell.value, datemode)
            return int(time.mktime(dt.timetuple()))
        elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
            return 1 if cell.value else 0
        else:
            raise Exception("Error: invalid cell type. type:{}".format(cell.ctype))

    def _getCellFloat(self, cell):
        if cell.ctype == xlrd.XL_CELL_EMPTY:
            return float(0.0)
        elif cell.ctype == xlrd.XL_CELL_TEXT:
            return float(cell.value)
        elif cell.ctype == xlrd.XL_CELL_NUMBER:
            return float(cell.value)
        elif cell.ctype == xlrd.XL_CELL_DATE:
            dt = xlrd.xldate.xldate_as_datetime(cell.value, datemode)
            return "%.2f" % time.mktime(dt.timetuple())
        elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
            return  1 if cell.value else 0
        else:
            raise Exception("Error: invalid cell type. type:{}".format(cell.ctype))

    def _getCellBool(self, cell):
        if cell.ctype == xlrd.XL_CELL_EMPTY:
            return False
        elif cell.ctype == xlrd.XL_CELL_TEXT:
            value = cell.value.lower()
            if value in BOOL_NO:
                return False
            elif value in BOOL_YES: 
                return True
            else:
                raise Exception("Error: can't switch the value to bool. value:{}".format(value))
        elif cell.ctype == xlrd.XL_CELL_NUMBER:
            return float(cell.value) != 0
        elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
            return cell.value
        else:
            raise Exception("Error: invalid cell type. type:{}".format(cell.ctype))

if __name__ == "__main__":
    parser = argparse.ArgumentParser("excel to lua converter")
    parser.add_argument("-i", dest="input_dir", help="excel表文件目录", default="../xls")
    parser.add_argument("-o", dest="output_dir", help="输出目录", default="../output")
    parser.add_argument("-t", dest="type", metavar='lua|json|all', help="导出类型(默认为导出为lua文件)", default="lua")
    parser.add_argument("-f", dest="force", help="强制导出所有表格", action="store_true")
    parser.add_argument("-e", dest="export", metavar="server|client", help="导出表格的类型（服务端导表or客户端导表）", default="server")
    args = parser.parse_args()
    converter = Converter(args, Logger())
    converter.convertAll()
