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
配置好excel表格之后，windows下直接启动xlsTools.bat开始一键导表（linux下启动脚本为xlsTools.sh），该脚本会自动将xls目录下的所有excel表格进行导出，导出目录存放在output目录中，默认导出为lua文件。如果需要修改默认导出规则，可执行scripts/xlsTools.py脚本导出，脚本参数如下：
```shell
# ./xlsTools.py -h
usage: excel to lua converter [-h] [-i INPUT_DIR] [-o OUTPUT_DIR] [-f] [-t TYPE]

options:
  -h, --help            show this help message and exit
  -i INPUT_DIR, --input_dir INPUT_DIR
                        excel表文件目录
  -o OUTPUT_DIR, --output_dir OUTPUT_DIR
                        输出目录
  -f, --force           强制导出所有表格
  -t TYPE, --type TYPE  导出类型
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
        "name": "\u6740\u6b7b10\u4e2a\u91ce\u602a",
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
        "name": "\u7d2f\u8ba1\u5b58\u6d3b50\u5206\u949f",
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