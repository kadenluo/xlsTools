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
from enum import Enum

datemode = 0 # 时间戳模式 0: 1900-based, 1: 1904-based
BOOL_YES = ["yes", "1", "是"]
BOOL_NO = ["", "nil", "0", "false", "no", "none", "否", "无"]
EXPORT_TYPES = ["lua", "json"]

class KeyType(Enum):
    FirstIndex = 1
    SecondIndex = 2
    Field = 3

#class ElementType(Enum):
#    ALL = "all"
#    CLIENT = "client"
#    SERVER = "server"

def myassert(expr, errmsg="Unknown"):
    if not expr:
        raise Exception(errmsg)

class Logger():
    def __init__(self):
        pass

    def debug(self, pattern, *args):
        print("[DEBUG] {}".format(pattern.format(*args)))

    def info(self, pattern, *args):
        print("[INFO] {}".format(pattern.format(*args)))

    def error(self, pattern, *args):
        print("[ERROR] {}".format(pattern.format(*args)))

    def critical(self, pattern, *args):
        print("[Critial] {}".format(pattern.format(*args)))

class Converter:
    _config = {} # 输入配置
    _indent = "    " #缩进
    _cachefile = "./.cache"
    _logger = None
    _sheetDict = {}
    def __init__(self, config, logger):
        self._logger = logger
        myassert(config.client_type is None or config.client_type == "all" or config.client_type in EXPORT_TYPES)
        myassert(config.server_type is None or config.server_type == "all" or config.server_type in EXPORT_TYPES)
        if not config.client_type is None:
            myassert(not config.client_output_dir is None)
        if not config.server_type is None:
            myassert((not config.server_output_dir is None))

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
            if data == "nil":
                return "nil"
            else:
                return '[[%s]]'%(data)
        elif isinstance(data, bool):
            return 'true' if data else 'false'
        else:
            return str(data)
        return ", \n".join(lines)

    def saveData(self, output_dir, filename, ftype, data):
        if ftype == "all" or ftype == "lua":
            luaTable = self._toLua(data)
            code = "local data = %s\n\nreturn data" % (luaTable)
            ast.parse(code)
            filepath = os.path.join(output_dir, "{}.lua".format(filename))
            self.saveFile(filepath, code)

        if ftype == "all" or ftype == "json":
            code = json.dumps(data, indent=4, ensure_ascii=False)
            filepath = os.path.join(output_dir, "{}.json".format(filename))
            self.saveFile(filepath, code)

    def saveFile(self, filepath, data):
        out_dir = os.path.dirname(filepath)
        if not os.path.exists(out_dir):
            os.makedirs(out_dir, mode=0o755)
        with open(filepath, 'wb') as f:
            f.write(data.encode('utf-8'))

    def convertAll(self):
        try:
            history = {}
            if not self._config.force:
                if os.path.exists(self._cachefile):
                    with open(self._cachefile, 'r', encoding='utf-8') as f:
                        history = json.load(f)

            allfiles = {}
            for filename in os.listdir(self._config.input_dir):
                if filename.startswith("~"):
                    continue

                filepath = os.path.join(self._config.input_dir, filename)
                mtime = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(os.stat(filepath).st_mtime)) 
                allfiles[filename] = mtime

                isInvalid = True
                for pattern in self._config.exclude_files:
                    if len(pattern)>0 and re.match(pattern, filename):
                        isInvalid = False
                        break

                if not isInvalid:
                    self._logger.info("convert {} but is excluded.", filename)
                    continue

                if ((filename not in history) or (history[filename] != mtime)):
                    self.convertFile(filename)
                    history[filename] = mtime
                    with open(self._cachefile, "w", encoding='utf-8') as f:
                        json.dump(history, f, indent=4, ensure_ascii=False)

            # 清理cache
            delkeys = []
            for (filename, mtime) in history.items():
                if filename not in allfiles:
                    delkeys.append(filename)
            for k in delkeys:
                del history[filename]
            with open(self._cachefile, "w", encoding='utf-8') as f:
                json.dump(history, f, indent=4, ensure_ascii=False)

            self._logger.info("success!!!")
        except Exception as ex:
            self._logger.error(traceback.format_exc())
            sys.exit(-1)

    def convertFile(self, filename):
        filepath = os.path.join(self._config.input_dir, filename)
        wb = xlrd.open_workbook(filepath)
        for sheet in wb.sheets():
            self._logger.info("convert {}({})...", filename, sheet.name)
            client, server = self._convertSheet(sheet)
            #print(json.dumps(client, indent=4))
            #print(json.dumps(server, indent=4))
            if not self._config.client_type is None and len(client) > 0:
                self.saveData(self._config.client_output_dir, sheet.name, self._config.client_type, client)
            if not self._config.server_type is None and len(server) > 0:
                self.saveData(self._config.server_output_dir, sheet.name, self._config.server_type, server)

    def _convertHead(self, sheet):
        nrows = sheet.nrows
        ncols = sheet.ncols

        clientDict = {"col2fields":{},"fields":{}, "indexes":{}}
        serverDict = {"col2fields":{},"fields":{}, "indexes":{}}
        for col in range(ncols):
            desc = self._getCellString(sheet.cell(0, col)).strip(' ')
            name = self._getCellString(sheet.cell(1, col)).strip(' ')
            vtype = self._getCellString(sheet.cell(2, col)).strip(' ')
            etype = self._getCellString(sheet.cell(3, col)).strip(' ').lower()

            if len(name) == 0:
                continue

            if etype == "":
                etype = "all"

            myassert(etype == "all" or etype == "client" or etype == "server", "etype is invalid. (etype=%s)"%etype)

            keyType = KeyType.Field
            isIndexCanRepeat = False
            if name.startswith('*'):
                keyType = KeyType.FirstIndex
                if name.endswith('*'):
                    isIndexCanRepeat = True
                if name.startswith('**'):
                    keyType = KeyType.SecondIndex
                name = name.strip('*')

            field = {"desc":desc, "name":name, "type":vtype, "levels":name.split('#')}
            if etype == "all" or etype == "client":
                if keyType == KeyType.FirstIndex or keyType == KeyType.SecondIndex:
                    clientDict["indexes"][keyType] = {"isCanRepeat":isIndexCanRepeat, "name":name}
                myassert(name not in clientDict["fields"], "field is repeated. (name=%s)"%name)
                clientDict["fields"][name] = col
                clientDict["col2fields"][col] = field

            if etype == "all" or etype == "server":
                if keyType == KeyType.FirstIndex or keyType == KeyType.SecondIndex:
                    serverDict["indexes"][keyType] = {"isCanRepeat":isIndexCanRepeat, "name":name}
                myassert(name not in serverDict["fields"], "field is repeated. (name=%s)"%name)
                serverDict["fields"][name] = col
                serverDict["col2fields"][col] = field

        #self._sheetDict[sheet.name] = {"client":clientDict, "server":serverDict}
        return clientDict, serverDict

    def _convertSheet(self, sheet):
        nrows = sheet.nrows
        ncols = sheet.ncols
        myassert ((nrows > 3) and (ncols > 1), "nrows(%d) or ncols(%d) is invalid." %(nrows, ncols))
        clientDict, serverDict = self._convertHead(sheet)

        if len(clientDict["indexes"]) > 0:
            clientResult = {}
        else:
            clientResult = []

        if len(serverDict["indexes"]) > 0:
            serverResult = {}
        else:
            serverResult = []

        #数据从第4行开始
        for row in range(4, nrows):
            clientFields = {}
            serverFields = {}
            clientItem = {"_meta":{"isdict":True}}
            serverItem = {"_meta":{"isdict":True}}
            for col in range(ncols):
                if sheet.cell(row, col).ctype == xlrd.XL_CELL_EMPTY:
                    continue
                if col in clientDict["col2fields"]:
                    meta = clientDict["col2fields"][col]
                    value = self._getCellValue(sheet.cell(row, col), meta["type"]) 
                    key = meta["name"]
                    if len(meta["levels"]) == 1:
                        clientItem[key] = value
                    else:
                        clientFields[key] = value

                if col in serverDict["col2fields"]:
                    meta = serverDict["col2fields"][col]
                    value = self._getCellValue(sheet.cell(row, col), meta["type"]) 
                    key = meta["name"]
                    if len(meta["levels"]) == 1:
                        serverItem[key] = value
                    else:
                        serverFields[key] = value

            # client
            if len(clientDict["col2fields"]) > 0 :
                self._genRowData(clientResult, clientDict, clientFields, clientItem)

            # server
            if len(serverDict["col2fields"]) > 0:
                self._genRowData(serverResult, serverDict, serverFields, serverItem)

        return clientResult, serverResult

    def _genRowData(self, result, sheetDict, fields, items):
        self._convertRow(items, fields)
        items = self._fixLevelType(items)
        if isinstance(fields, list):
            result.append(items)
        else:
            indexNum = len(sheetDict["indexes"])
            for idx in range(1, indexNum+1):
                myassert(KeyType(idx) in sheetDict["indexes"], "缺少%d级索引!"%(idx))

            if indexNum == 1:
                index = sheetDict["indexes"][KeyType.FirstIndex]
                k = items[index["name"]]
                if len(items) == 2:
                    myassert(not k in result, "primary must not be repeated!(key=%s)"%(str(k)))
                    del items[index["name"]]
                    it = list(items.items())[0]
                    result[k] = it[1]
                else:
                    if not index["isCanRepeat"]:
                        myassert(k not in result, "一级索引重复！key:%s"%(k))
                        result[k] = items
                    else:
                        if k not in result:
                            result[k] = []
                        result[k].append(items)

            elif indexNum == 2:
                firstIndex = sheetDict["indexes"][KeyType.FirstIndex]
                secondIndex = sheetDict["indexes"][KeyType.SecondIndex]
                firstIndexName = firstIndex["name"]
                secondIndexName = secondIndex["name"]

                firstIndexValue = items[firstIndexName] 
                secondIndexValue = items[secondIndexName] 

                del items[firstIndexName]
                del items[secondIndexName]

                if firstIndexValue not in result:
                    result[firstIndexValue] = {}

                myassert(secondIndexValue not in result[firstIndexValue], "二级键重复！(subKey=%s), (mainKey=%s)"%(str(secondIndexValue), str(firstIndexValue)))
                result[firstIndexValue][secondIndexValue] = items

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
            bAllNumber = True
            for k in item.keys():
                if not str.isdigit(k):
                    bAllNumber = False
                    break
            if bAllNumber:
                for k in sorted(item.keys(), key=lambda key: int(key)):
                    tmp.append(item[k])
            else:
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
            if cell.value % 1 == 0 :
                return str(int(cell.value))
            else:
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
    parser.add_argument("--input_dir", dest="input_dir", help="excel表文件目录", default="../xls")
    parser.add_argument("--client_type", dest="client_type", metavar='|'.join(["all"]+EXPORT_TYPES), help="客户端导出类型(默认为导出为lua文件)", default="lua")
    parser.add_argument("--client_output_dir", dest="client_output_dir", help="client输出目录", default="../output/client")
    parser.add_argument("--server_type", dest="server_type", metavar='|'.join(["all"]+EXPORT_TYPES), help="服务端导出类型(默认为导出为lua文件)", default="lua")
    parser.add_argument("--server_output_dir", dest="server_output_dir", help="server输出目录", default="../output/server")
    parser.add_argument("--force", dest="force", help="强制导出所有表格", action="store_true")
    parser.add_argument("--exclude_files", dest="exclude_files", help="排除文件", type=str, nargs="+", default=[])
    args = parser.parse_args()
    converter = Converter(args, Logger())
    converter.convertAll()
