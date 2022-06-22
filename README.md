# xlsTools
xlsTools 用于将策划配置的excel表格转换成lua或json文件，支持xls和xlsx格式。
# 使用示例
目录结构
```lua
|--scripts
|  |--xlsTools.py  # excel转表工具
|--xls   # excel目录，以下是两个示例文件
|  |--任务表.xlsx  
|  |--系统参数表.xlsx
|--xlsTools.bat # windows下的程序入口
|--doc # 文档
```
## sheet配置
导表时以sheet为单位导出，一个sheet导出一个lua或json文件（sheet名即为导出文件的文件名，比如：sheet名为hello，导出之后的文件名为hello.lua或hello.json），一个excel里可以存在多个sheet。
## 数据配置
使用时，只需按照以下格式配置excel即可，表的前3行为表结构描述，表的数据从第4行开始。  
第1行：字段描述，仅用于方便查看，不导出到输出文件中；  
第2行：字段名称，字段名称遵循一下规则：  
1. 带星号的字段为主key，一个sheet里最多只能有一个主key， 也可以没有；
2. 多层结构可通过#符号来表示，比如：“awards#1#id” 表示三层结构，三层结构的key分别为award，1，id。 如果key为数字，则该层为数组类型，如果key不是数字，则该层为字典类型。  

第3行：字段类型，只支持4种基础字段类型：int， float，bool,  string。可通过字段名称用基础类型组合成复杂的嵌套类型，比如数组、字典类型;  
第4+行: 具体数据值。  
![示例配置](/doc/images/任务表.png)  

## 导表工具
配置好excel表格之后，启动根目录下的xlsToolsGUI.exe工具，启动界面如下：
![gui](/doc/images/gui.png)
在linux端，可能没有GUI，此时可以直接调用转表脚本（./scripts/xlsTools.py）进行转表，使用如下：

```shell
# ./xlsTools.py -h
usage: excel to lua converter [-h] [-i INPUT_DIR] [-c CLIENT_OUTPUT_DIR] [-s SERVER_OUTPUT_DIR] [-t lua|json|all] [-f] [-e EXCLUDE_FILES [EXCLUDE_FILES ...]]

options:
  -h, --help            show this help message and exit
  -i INPUT_DIR          excel表文件目录
  -c CLIENT_OUTPUT_DIR  client输出目录
  -s SERVER_OUTPUT_DIR  server输出目录
  -t lua|json|all       导出类型(默认为导出为lua文件)
  -f                    强制导出所有表格
  -e EXCLUDE_FILES [EXCLUDE_FILES ...]
                        排除文件(正则匹配)
```
## 输出文件
### lua
```lua
return {
    [1000] = {
        id = 1000, 
        type = 1, 
        name = [[杀死10个野怪]], 
        start_time = 1577836800, 
        end_time = 1609459200, 
        conds = {
            10, 
            0, 
            0
        }, 
        awards = {
            {
                expire = 0, 
                id = 1000, 
                num = 100
            }, 
            {
                expire = 0, 
                id = 1001, 
                num = 100
            }, 
            {
                expire = 0, 
                id = 0, 
                num = 0
            }
        }
    }, 
    [1001] = {
        id = 1001, 
        type = 2, 
        name = [[累计存活50分钟]], 
        start_time = 1575158400, 
        end_time = 1577836800, 
        conds = {
            50, 
            0, 
            0
        }, 
        awards = {
            {
                expire = 0, 
                id = 1000, 
                num = 100
            }, 
            {
                expire = 0, 
                id = 1001, 
                num = 100
            }, 
            {
                expire = 0, 
                id = 0, 
                num = 0
            }
        }
    }
}
```
### json
```json
{
    "1000": {
        "id": 1000,
        "type": 1,
        "name": "杀死10个野怪",
        "start_time": 1577874030,
        "end_time": 1609459200,
        "conds": [
            true,
            false,
            false
        ],
        "awards": [
            {
                "expire": 0,
                "id": 1000,
                "num": 100
            },
            {
                "expire": 0,
                "id": 1001,
                "num": 100
            },
            {
                "expire": 0,
                "id": 0,
                "num": 0
            }
        ]
    },
    "1001": {
        "id": 1001,
        "type": 2,
        "name": "累计存活50分钟",
        "start_time": 1577874030,
        "end_time": 1577836800,
        "conds": [
            true,
            false,
            false
        ],
        "awards": [
            {
                "expire": 0,
                "id": 1000,
                "num": 100
            },
            {
                "expire": 0,
                "id": 1001,
                "num": 100
            },
            {
                "expire": 0,
                "id": 0,
                "num": 0
            }
        ]
    }
}
```

# TODO
* 枚举变量的支持；
* pb导出；