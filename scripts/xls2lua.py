#!/usr/bin/python2
# -*-  coding:utf-8 -*-

import os
import re
import sys
import time
import datetime
import xlrd
import argparse

sys.path.append("..")

from protocol import xls_pb2

datemode = 0 # 时间戳模式 0: 1900-based, 1: 1904-based

class Converter:
    _config = {}
    _xls2class = {}
    _indent = u"    "
    def __init__(self, config):
        self._config = config
        for line in open(self._config.maps):
            items = line.split('=')
            filename = items[0].strip().decode('utf-8')
            classname = items[1].strip(u" \n").decode('utf-8')
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
                if isinstance(key, str) or isinstance(key, unicode):
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
            for key in sorted(data.keys(), func):
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
            if not mainkey:
                rows.append(u"%s%s" % (self._indent, tostring(row, func=field_sort_func)))
            else:
                key = row[mainkey]
                del row[mainkey]
                left_split = u"["
                right_split = u"]"
                if isinstance(key, str) or isinstance(key, unicode):
                    left_split = u""
                    right_split = u""
                    key = key.strip(u'"')
                if len(row.items()) == 1:
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
        if classname not in xls_pb2.DESCRIPTOR.message_types_by_name:
            raise Exception(u"Failed to find the class name. %s" % classname)
        desc = xls_pb2.DESCRIPTOR.message_types_by_name[classname]
        assert (desc.fields[0].label == 3) and (desc.fields[0].cpp_type == 10)
        desc = desc.fields_by_name['list'].message_type

        filepath = u"%s/%s" % (self._config.input_dir, filename)
        worksheet = xlrd.open_workbook(filepath)
        sheet = worksheet.sheet_by_index(0)
        nrows = sheet.nrows
        ncols = sheet.ncols
        assert ((nrows > 2) and (ncols > 1))
        field2index = {}
        mainkey = None
        for col in xrange(ncols):
            value = sheet.cell(1, col).value
            if value.endswith('*'):
                value = value.strip('*')
                assert(mainkey == None)
                mainkey = value
            field2index[value] = col

        result = []
        for i in xrange(2, nrows):
            item = self._convertRow(desc, field2index, sheet, i, "")
            result.append(item)

        def field_sort_func(x, y):
            if x in field2index and y in field2index:
                return field2index[x] - field2index[y]
            elif x in field2index and y not in field2index:
                return -1
            elif x not in field2index and y in field2index:
                return 1
            else:
                return x < y
        code = self.getCode(result, mainkey, field_sort_func)
        code = u"_G.tables = _G.tables or {}\n_G.tables.%s = %s" % (classname, code)
        print(code)

        # save
        self.save(classname + '.lua', code)

    def _convertRow(self, desc, field2index, sheet, row, prefix):
        row_content = {}
        for field in desc.fields:
            item = None
            if field.cpp_type == 10:
                child_desc = desc.fields_by_name[field.name].message_type
                if field.label != 3:
                    child_prefix = u"%s%s_" % (prefix, field.name)
                    item = self._convertRow(child_desc, field2index, sheet, row, child_prefix)
                else:
                    item = []
                    for idx in xrange(1, 10):
                        child_prefix = u"%s%s_%d_" % (prefix, field.name, idx)
                        ishave = False
                        for key in field2index.keys():
                            if re.match(child_prefix, key):
                                ishave = True
                                break
                        if ishave == False:
                            break
                        it = self._convertRow(child_desc, field2index, sheet, row, child_prefix)
                        item.append(it)
            elif field.label == 3:
                item = []
                for idx in xrange(1, 20):
                    name = u"%s%s_%d" % (prefix, field.name, idx)
                    if name not in field2index:
                        break
                    cell = sheet.cell(row, field2index[name])
                    item.append(self._getCellValue(cell, field.cpp_type))
            else:
                name = u"%s%s" % (prefix, field.name)
                if name not in field2index:
                    raise Exception(u"Can't find the field. %s" % name)
                cell = sheet.cell(row, field2index[name])
                item = self._getCellValue(cell, field.cpp_type)
            row_content[field.name] = item
        return row_content

    def _getCellValue(self, cell, field_type):
        if field_type >= 1 and field_type <= 4: # int32 uint32 int64 uint64
            return self._getCellInt(cell)
        elif field_type == 5 or field_type == 6: # double float
            return self._getCellFloat(cell)
        elif field_type == 7: # bool
            return self._getCellBool(cell)
        elif field_type == 9: # string
            return self._getCellString(cell)
        else:
            raise Exception(u"This type is invalid. %s" % field_type)
        
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
        return u'"%s"' % cell_text
    
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
