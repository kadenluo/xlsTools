#!/usr/bin/python3
# -*-  coding:utf-8 -*-
import time
import xlrd
from utils import *

datemode = 0 # 时间戳模式 0: 1900-based, 1: 1904-based
BOOL_YES = ["yes", "1", "是"]
BOOL_NO = ["", "nil", "0", "false", "no", "none", "否", "无"]

#singleton
class DataParser():
    def __init__(self):
        pass

    def getCellValue(self, cell, vtype):
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
        elif cell.ctype == xlrd.XL_CELL_EMPTY:
            return []
        else:
            print("======", cell.ctype)
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
        elif cell.ctype == xlrd.XL_CELL_EMPTY:
            return []
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
        elif cell.ctype == xlrd.XL_CELL_EMPTY:
            return []
        else:
            raise Exception("Error: invalid cell type. type:{}".format(cell.ctype))

    def _getCellListForString(self, cell):
        if cell.ctype == xlrd.XL_CELL_TEXT:
            value = cell.value.strip(' ').lstrip('[').rstrip(']')
            if len(value) == 0:
                return []
            return value.split(',')
        elif cell.ctype == xlrd.XL_CELL_EMPTY:
            return []
        else:
            raise Exception("Error: invalid cell type. type:{}".format(cell.ctype))
        
    def _getCellString(self, cell):
        if cell.ctype == xlrd.XL_CELL_EMPTY:
            return ""
        elif cell.ctype == xlrd.XL_CELL_TEXT:
            return cell.value.strip()
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
