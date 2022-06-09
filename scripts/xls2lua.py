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

class Converter:
    _config = {}
    _xls2class = {}
    _indent = u"    "
    def __init__(self, config):
        self._config = config
        for line in open(self._config.maps):
            items = line.split('=')
            filename = items[0].strip().encode("utf-8").decode('utf-8')
            classname = items[1].strip(u" \n").encode("utf-8").decode('utf-8')
            self._xls2class[filename] = classname

    def _objectToStringWithIndent(self, data, level=1, func=None):
        if isinstance(data, list) or isinstance(data, tuple):
            content = []
            for item in data:
                content.append(self._objectToStringWithIndent(item, level + 1, func))
            content = self._indent*(level+1) + (u",\n"+self._indent*(level+1)).join(content)
            result = u"{\n%s\n%s}" % (content, self._indent*(level))
        elif isinstance(data, dict):
            content = []
            for key in sorted(data.keys(), func):
                value = data[key]
                if isinstance(key, str):
                    content.append(u"%s = %s" % (key.strip('"'), self._objectToStringWithIndent(value, level+1, func)))
                else:
                    content.append(u"[%s] = %s" % (key, self._objectToStringWithIndent(value, level+1, func)))
            content = self._indent*(level+1) + (u",\n"+self._indent*(level+1)).join(content)
            result = u"{\n%s\n%s}" % (content, self._indent*level)
        else:
            result = u"%s" % data
        return result

    def _objectToString(self, data, func=None):
        if isinstance(data, list) or isinstance(data, tuple):
            content = []
            for item in data:
                content.append(self._objectToString(item, func))
            content = u",".join(content)
            result = u"{%s}" % (content)
        elif isinstance(data, dict):
            content = []
            for key in sorted(data.keys(), key=func):
                value = data[key]
                content.append(u"%s=%s" % (key, self._objectToString(value, func)))
            content = u",".join(content)
            result = u"{%s}" % (content)
        else:
            result = u"%s" % data
        return result

    def getCode(self, data, mainkey, field_sort_func=None):
        rows = []
        tostring = self._objectToString
        for row in data:
            print("============", data)
            if not mainkey:
                rows.append(u"%s%s" % (self._indent, tostring(row, func=field_sort_func)))
            else:
                key = row[mainkey]
                del row[mainkey]
                left_split = u"["
                right_split = u"]"
                if isinstance(key, str):
                    left_split = u""
                    right_split = u""
                    key = key.strip(u'"')
                if len(row.items()) == 1:
                    print("============", data)
                    rows.append(u"%s%s%s%s = %s" % (self._indent, left_split, key,  right_split, tostring(row.items()[0][1], func=field_sort_func)))
                else:
                    rows.append(u"%s%s%s%s = %s" % (self._indent, left_split, key, right_split, tostring(row, func=field_sort_func)))
        code = u"{\n%s\n}" % (u",\n".join(rows))
        return code

    def save(self, filename, data):
        lua_dir = self._config.output_dir
        if not os.path.exists(lua_dir):
            os.makedirs(lua_dir)
        filepath = u"%s/%s"% (lua_dir, filename)
        with open(filepath, 'wb') as f:
            f.write(data.encode('utf-8'))

    def convertAll(self):
        for filename in self._xls2class.keys():
            self.convert(filename)

    def convert(self, filename):
        if filename not in self._xls2class:
            raise Exception(u"Failed to load config of this file. %s" % filename)
        classname = self._xls2class[filename]

        filepath = u"%s/%s" % (self._config.input_dir, filename)
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
        #print(json.dumps(field2index, indent=4))

        result = []
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
            #print(json.dumps(item, indent=4))
            item = self._fixLevelType(item)
            #print(json.dumps(item, indent=4))
            result.append(item)

        #print(result)
        print(json.dumps(result, indent=4))

        #def field_sort_func(x, y):
        #    if x in field2index and y in field2index:
        #        return field2index[x] - field2index[y]
        #    elif x in field2index and y not in field2index:
        #        return -1
        #    elif x not in field2index and y in field2index:
        #        return 1
        #    else:
        #        return x < y

        #def field_sort_func(x):
        #    if x in field2index:
        #        return field2index[x]
        #    else:
        #        return 0
        #code = self.getCode(result, mainkey, field_sort_func)
        #code = u"_G.tables = _G.tables or {}\n_G.tables.%s = %s" % (classname, code)
        #print(code)

        # save
        #self.save(classname + '.lua', code)

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
            item = list(item.values()) # todo 检查values返回的顺序
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
            raise Exception(u"This type is invalid. %s" % vtype)
        
    def _getCellString(self, cell):
        cell_text = ""
        if cell.ctype == xlrd.XL_CELL_TEXT:
            cell_text = cell.value
        elif cell.ctype == xlrd.XL_CELL_NUMBER:
            cell_text = (u"%.2f" % cell.value).rstrip('0').rstrip('.')
        elif cell.ctype == xlrd.XL_CELL_DATE:
            dt = xlrd.xldate.xldate_as_datetime(cell.value, datemode)
            cell_text = u"%s" % dt
        elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
            cell_text =  u"true" if cell.value else u"false"
        return cell_text
    
    def _getCellInt(self, cell):
        if cell.ctype == xlrd.XL_CELL_TEXT or cell.ctype == xlrd.XL_CELL_NUMBER:
            return int(cell.value)
        elif cell.ctype == xlrd.XL_CELL_DATE:
            dt = xlrd.xldate.xldate_as_datetime(cell.value, datemode)
            return int(time.mktime(dt.timetuple()))
        elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
            return 1 if cell.value else 0
        return 0

    def _getCellFloat(self, cell):
        if cell.ctype == xlrd.XL_CELL_TEXT:
            return float(cell.value)
        elif cell.ctype == xlrd.XL_CELL_NUMBER:
            return float((u"%.2f" % cell.value).rstrip('0').rstrip('.'))
        elif cell.ctype == xlrd.XL_CELL_DATE:
            dt = xlrd.xldate.xldate_as_datetime(cell.value, datemode)
            return u"%.2f" % time.mktime(dt.timetuple())
        elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
            return  1 if cell.value else 0
        return 0.0

    def _getCellBool(self, cell):
        text = cell.value.lower()
        if text in [u"", u"nil", u"0", u"false", u"no", u"none", u"否", u"无"]:
            return u"false"
        return u"true"

if __name__ == "__main__":
    parser = argparse.ArgumentParser("excel to lua converter")
    parser.add_argument(u"-i", u"--input_dir", dest=u"input_dir", help=u"xls dirname", default=u"../xls")
    parser.add_argument(u"-o", u"--output_dir", dest=u"output_dir", help=u"outpute dirname", default=u"../output/lua")
    parser.add_argument(u"-m", u"--map_file", dest=u"maps", help=u"xls to struct map", default=u"../config.conf")
    args = parser.parse_args()
    converter = Converter(args)
    converter.convertAll()
