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
from operator import itemgetter

sys.path.append("..")

from google.protobuf.descriptor import FieldDescriptor

datemode = 0 # 时间戳模式 0: 1900-based, 1: 1904-based
BOOL_YES = ["yes", "1", "是"]
BOOL_NO = ["", "nil", "0", "false", "no", "none", "否", "无"]

class Converter:
    _config = {}
    _xls2class = {}
    _indent = "    "
    def __init__(self, config):
        self._config = config
        for line in open(self._config.maps):
            items = line.split('=')
            filename = items[0].strip().encode("utf-8").decode('utf-8')
            classname = items[1].strip(" \n").encode("utf-8").decode('utf-8')
            self._xls2class[filename] = classname

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
        out_dir = "{}/{}".format(self._config.output_dir, output_type)
        if not os.path.exists(out_dir):
            os.makedirs(out_dir)
        filepath = "{}/{}.{}".format(out_dir, filename, output_type)
        with open(filepath, 'wb') as f:
            f.write(data.encode('utf-8'))

    def convertAll(self):
        for filename in self._xls2class.keys():
            self.convert(filename)

    def convert(self, filename):
        if filename not in self._xls2class:
            raise Exception("Failed to load config of this file. %s" % filename)
        classname = self._xls2class[filename]

        filepath = "%s/%s" % (self._config.input_dir, filename)
        worksheet = xlrd.open_workbook(filepath)
        sheet = worksheet.sheet_by_index(0)
        nrows = sheet.nrows
        ncols = sheet.ncols
        assert ((nrows > 2) and (ncols > 1))

        mainkey = None
        field2index = {}
        for col in range(ncols):
            desc = self._getCellString(sheet.cell(0, col)).strip(' ')
            name = self._getCellString(sheet.cell(1, col)).strip(' ')
            vtype =self._getCellString(sheet.cell(2, col)).strip(' ')

            if len(name) == 0:
                continue

            if name.startswith('*'):
                name = name.strip('*')
                assert(mainkey == None)
                mainkey = name
            field2index[col] = {"desc":desc, "name":name, "type":vtype, "levels":name.split('#')}

        if mainkey is None:
            result = []
        else:
            result = {}
        for row in range(3, nrows):
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
                result[k] = item

        #print(json.dumps(result, indent=4))

        if self._config.type == "lua":
            luaTable = self._toLua(result)
            code = "return %s" % (luaTable)
        elif self._config.type == "json":
            code = json.dumps(result, indent=4)
        else:
            raise Exception("Error: invalid output type.")

        # save
        self.save(self._config.type, classname, code)

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
        elif vtype == "float" or vtype == "double":
            return self._getCellFloat(cell)
        elif vtype == "bool":
            return self._getCellBool(cell)
        elif vtype == "string": # string
            return self._getCellString(cell)
        else:
            raise Exception("This type is invalid. %s" % vtype)
        
    def _getCellString(self, cell):
        if cell.ctype == xlrd.XL_CELL_EMPTY:
            return ""
        elif cell.ctype == xlrd.XL_CELL_TEXT:
            return cell.value
        elif cell.ctype == xlrd.XL_CELL_NUMBER:
            return ("%.2f" % cell.value).rstrip('0').rstrip('.')
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
            return float(("%.2f" % cell.value).rstrip('0').rstrip('.'))
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
    parser.add_argument("-i", "--input_dir", dest="input_dir", help="excel表文件目录", default="../xls")
    parser.add_argument("-o", "--output_dir", dest="output_dir", help="输出目录", default="../output")
    parser.add_argument("-m", "--map_file", dest="maps", help="excel表与输出文件之间映射", default="../config.conf")
    parser.add_argument("-t", "--type", dest="type", help="导出类型", default="lua")
    args = parser.parse_args()
    converter = Converter(args)
    converter.convertAll()
