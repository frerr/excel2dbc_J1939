# excel2dbc-py
**简介**

Inceptio自用DF excel格式转dbc，只针对DF excel模板，生成对应J1939 CAN DBC文件

**版本说明**

- v1：

  - 初版释放

  - 报错信息处理待完善

**EXCEL格式要求**

- 模板格式参见template下DF_Example.xlsx

- **重要1**：每一个sheet里，ECU节点命名，禁止出现"/"和空格字符如“ABS/EPS”或“ABS EPS”，否则Vector CANdbc工具报错。可连续如“ABSEPS”

- **重要2**：每一个sheet里，Signal name命名，禁止出现“/”和空格字符如“vehicle/speed”或“vehicle speed”，否则Vector CANdbc工具报错。推荐使用“_”如“vehicle_speed”

- **重要3**：每一个sheet里，Signal value定义，编写格式强制要求如下：

  0x0:Trailer ABS not fully operational
  0x1:Trailer ABS fully operational
  0x2:Error
  0x3:Not available

  其中“：”可替代成“=”，空格或换行符不敏感
  
- **重要4**：每一个sheet里，Message cycle time只可出现数值，如“200”

