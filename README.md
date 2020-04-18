# xlsTools
xlsTools 用于将策划配置的excel表格转换成代码配置文件，比如将“.xls”、“.xlsx”格式的文件转换lua配置文件（目前只支持转换成lua文件）。
# 使用示例
目录结构
```lua
|--protocol  
|  |--xls.proto  # protobuf 协议文件配置
|--scripts
|  |--xls2lua.py  # excel转lua的转换工具
|--tools
|  |--protoc.exe  # 转换protobuf协议的工具
|--xls   # excel目录，以下是两个示例文件
|  |--任务表.xlsx  
|  |--系统参数表.xlsx
|--config.conf # 转换配置文件
|--xlsTools.bat # windows下的程序入口
|--doc # 文档
```
使用的时候需要修改3个地方：config.conf、protocol/xls.proto、xls/xxx.xlsx。具体使用步骤如下：
1. 在protocol/xls.proto中添加新的数据结构定义：
```lua
syntax = "proto2";

message AwardItem{
    required int32 id = 1;
    required int32 num = 2;
    optional int32 expire = 3;
}

message TaskItem{
    required int32 id = 1;
    required int32 type = 2;
    required string name = 3;
    required uint32 start_time = 4;
    required uint32 end_time = 5;
    repeated int32  conds = 6;
    repeated AwardItem awards = 7;
}

message TaskList{
    repeated TaskItem list = 1; 
}
```
2. 在xls目录下新增excel文件；  
![示例配置](/doc/任务表.png)  
如图，表格的第一行是标题，可以随便自定义。第二个行是结构定义，必须和protocol/xls.proto中定义的结构体对应。其规则如下：  
a. 带“*”的表示这一列为主键，其对应的结构里不带“*”，一个表格里最多只能有一个主键；  
b. 普通类型（即非repeated且类型为非结构类型）保持excel和结构体名对应即可，如“任务表.xlsx”里的id，type，name，start_time，end_time这些字段；  
c. repeated类型的组合方式是:“结构体名_” + 索引。如“任务表.xlsx”里的cond_1, cond_2, cond_3；  
d. 结构类型的，比如TaskItem.awards。其对应的应该是“任务表.xlsx”里的“awards_” + 索引 + “_id”, “awards_” + 索引 + “_num”, “awards_” + 索引 + “_expire”，这里的索引是因为TaskItem.awards是repeated类型的。如果TaskItem.awards是非repeated类型的，则其对应便是“awards_id”, “awards_num”, “awards_expire”。  
3. 在config.conf目录下新增excel表格到具体结构的映射关系。如下：
```lua
任务表.xlsx = TaskList
系统参数.xlsx = SystemParams
```
4. 运行xlsTools.bat程序进行转换。转换结果在output/lua下。如：
```lua
_G.tables = _G.tables or {}
_G.tables.TaskList = {
    [1000] = {type=1,name="杀死10个野怪",start_time=1577808000,end_time=1609430400,conds={10,0,0},awards={id=1000,num=100,expire=0}},
    [1001] = {type=2,name="累计存活50分钟",start_time=1575129600,end_time=1577808000,conds={50,0,0},awards={id=1000,num=100,expire=0}}
}
```

# 注意事项
1. 目前支持将excel配置文件转换为lua文件，其它语言的支持待开发；
2. 转换的时候只会转换每个excel表格的第一个sheet，所以每个excel里最多只能有1个sheet；

