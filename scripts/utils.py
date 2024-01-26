#!/usr/bin/python3
# -*-  coding:utf-8 -*-

import os
import json
import hashlib
from luaparser import ast

output_indent = "    " #缩进

def singleton(cls):
    instances = {}
    def wrapper(*args, **kwargs):
        if cls not in instances:
            instances[cls] = cls(*args, **kwargs)
        return instances[cls]
    return wrapper

def getFileMD5(filepath):
    with open(filepath, 'rb') as fp:
        return hashlib.md5(fp.read()).hexdigest()
    return None

def toLua(data, level=1):
    lines = []
    if isinstance(data, list):
        items = []
        for value in data:
            value = toLua(value, level+1)
            items.append("%s%s"%(output_indent*level, value))
        lines.append("{\n%s\n%s}"%(", \n".join(items), output_indent*(level-1)))
    elif isinstance(data, dict):
        items = []
        keys = sorted(data.keys())
        for key in keys:
            value = data[key]
            if isinstance(key, int):
                key = "[%d]"%(key)
            elif isinstance(key, str):
                pass
            else:
                raise Exception("Error: {}({}) can't be a key.".format(key, type(key)))
            value = toLua(value, level+1)
            items.append("{}{} = {}".format(output_indent*level, key, value))
        lines.append("{\n%s\n%s}"%(", \n".join(items), output_indent*(level-1)))
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
    
def saveData(output_dir, filename, ftype, data):
    if ftype == "all" or ftype == "lua":
        luaTable = toLua(data)
        code = "local data = %s\n\nreturn data" % (luaTable)
        ast.parse(code)
        filepath = os.path.join(output_dir, "{}.lua".format(filename))

    if ftype == "all" or ftype == "json":
        code = json.dumps(data, indent=4, ensure_ascii=False, sort_keys=True)
        filepath = os.path.join(output_dir, "{}.json".format(filename))

    out_dir = os.path.dirname(filepath)
    if not os.path.exists(out_dir):
        os.makedirs(out_dir, mode=0o755)
    with open(filepath, 'wb') as f:
        f.write(code.encode('utf-8'))
