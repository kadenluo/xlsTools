#!/usr/bin/python3
# -*-  coding:utf-8 -*-

import os
import re
import xlrd
import argparse
import traceback
import utils
from enum import Enum
from logger import Logger
from dataParser import DataParser

MD5_DIR = "./.md5"
EXPORT_TYPES = ["lua", "json"]

class KeyType(Enum):
    FirstIndex = 1
    SecondIndex = 2
    Field = 3

def myassert(expr, errmsg="Unknown"):
    if not expr:
        raise Exception(errmsg)

def _fixLevelType(item):
    meta = item["_meta"]
    del item["_meta"]

    for (key, value) in item.items():
        if isinstance(value, dict):
            item[key] = _fixLevelType(value)

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


def _convertHead(sheet):
    nrows = sheet.nrows
    ncols = sheet.ncols

    clientDict = {"col2fields":{},"fields":{}, "indexes":{}}
    serverDict = {"col2fields":{},"fields":{}, "indexes":{}}
    for col in range(ncols):
        desc = DataParser().getCellValue(sheet.cell(0, col), "string")
        name = DataParser().getCellValue(sheet.cell(1, col), "string")
        vtype = DataParser().getCellValue(sheet.cell(2, col), "string")
        etype = DataParser().getCellValue(sheet.cell(3, col), "string").lower()

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

        name = name.strip()
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

    return clientDict, serverDict

def _convertSheet(sheet):
    nrows = sheet.nrows
    ncols = sheet.ncols
    myassert ((nrows > 3) and (ncols > 1), "nrows(%d) or ncols(%d) is invalid." %(nrows, ncols))
    clientDict, serverDict = _convertHead(sheet)

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
            #if sheet.cell(row, col).ctype == xlrd.XL_CELL_EMPTY:
            #    continue
            if col in clientDict["col2fields"]:
                meta = clientDict["col2fields"][col]
                value = DataParser().getCellValue(sheet.cell(row, col), meta["type"]) 
                key = meta["name"]
                if len(meta["levels"]) == 1:
                    clientItem[key] = value
                else:
                    clientFields[key] = value

            if col in serverDict["col2fields"]:
                meta = serverDict["col2fields"][col]
                value = DataParser().getCellValue(sheet.cell(row, col), meta["type"]) 
                key = meta["name"]
                if len(meta["levels"]) == 1:
                    serverItem[key] = value
                else:
                    serverFields[key] = value

        # client
        if len(clientDict["col2fields"]) > 0 :
            _genRowData(clientResult, clientDict, clientFields, clientItem, False)

        # server
        if len(serverDict["col2fields"]) > 0:
            _genRowData(serverResult, serverDict, serverFields, serverItem, True)

    return clientResult, serverResult

def _convertRow(result, fields):
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

    _convertRow(result, childfields)


def _genRowData(result, sheetDict, fields, items, flag):
    _convertRow(items, fields)
    items = _fixLevelType(items)
    if isinstance(fields, list):
        result.append(items)
    else:
        indexNum = len(sheetDict["indexes"])
        for idx in range(1, indexNum+1):
            myassert(KeyType(idx) in sheetDict["indexes"], "缺少%d级索引!"%(idx))

        if indexNum == 1:
            index = sheetDict["indexes"][KeyType.FirstIndex]
            k = items[index["name"]]
            if flag and len(items) == 2:
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


def convertFile(filepath, client_type, client_output_dir, server_type, server_output_dir):
    wb = xlrd.open_workbook(filepath)
    for sheet in wb.sheets():
        dirname, filename = os.path.split(filepath)
        if sheet.name.startswith("Sheet"):
            Logger().error("convert ({}:{}) failed with invalid sheet name.", filename, sheet.name)
            return False

        Logger().info("convert {} to {}...", filename, sheet.name)
        client, server = _convertSheet(sheet)
        if not client_type is None and len(client) > 0:
            utils.saveData(client_output_dir, sheet.name.lower(), client_type, client)
        if not server_type is None and len(server) > 0:
            utils.saveData(server_output_dir, sheet.name.lower(), server_type, server)

    if not os.path.exists(MD5_DIR):
        os.makedirs(MD5_DIR)
    with open(os.path.join(MD5_DIR, "{}.md5".format(os.path.basename(filepath))), "w") as f:
        f.write(utils.getFileMD5(filepath))

    return True

def convertFiles(files, client_type="json", client_output_dir="./output/client", server_type="json", server_output_dir="./output/server"):
    try:
        myassert(client_type is None or client_type == "all" or client_type in EXPORT_TYPES)
        myassert(server_type is None or server_type == "all" or server_type in EXPORT_TYPES)
        if not client_type is None:
            myassert(not client_output_dir is None)
        if not server_type is None:
            myassert((not server_output_dir is None))

        for filepath in files:
            ret = convertFile(filepath, client_type, client_output_dir, server_type, server_output_dir)
            if not ret:
                Logger().error("============FAILED============={}", filepath)
                return False
        
        Logger().info("============SUCCESS=============")
        return True
    except Exception as ex:
        Logger().error("convertFiles has error. traceback:{}", traceback.format_exc())
        return False
    
def getAllFiles(targetDir, excludeFiles=[".svn", ".git"]):
    allFiles = []
    for filename in os.listdir(targetDir):
        if os.path.basename(filename).startswith('~'):
            continue
        isInvalid = True
        for pattern in excludeFiles:
            if len(pattern)>0 and re.match(pattern, filename):
                isInvalid = False
                Logger().info("convert {} but is excluded. pattern:{}", filename, pattern)
                break

        if not isInvalid:
            continue
        
        filepath = os.path.join(targetDir, filename)
        allFiles.append(filepath)

    return allFiles

def getModifiedFiles(targetDir, md5Dir=MD5_DIR):
    modifiedFiles = []
    file2md5 = {}
    if os.path.exists(md5Dir):
        for filename in os.listdir(md5Dir):
            filepath = os.path.join(md5Dir, filename)
            filename, ext = os.path.splitext(filename)
            if ext != ".md5" or not os.path.exists(filepath):
                continue
            with open(filepath, "rb") as f:
                file2md5[os.path.join(targetDir, filename)] = f.read()
    
    for filename in os.listdir(targetDir):
        if filename.startswith("~"):
            continue
        filepath = os.path.join(targetDir, filename)
        md5 = utils.getFileMD5(filepath)
        if ((filepath not in file2md5) or (file2md5[filepath] != md5)):
            modifiedFiles.append(filepath)
    
    return modifiedFiles

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
    if not args.force:
        allFiles = getModifiedFiles(args.input_dir)
    else:
        allFiles = getAllFiles(args.input_dir, args.exclude_files)
    convertFiles(allFiles, args.client_type, args.client_output_dir, args.server_type, args.server_output_dir)
