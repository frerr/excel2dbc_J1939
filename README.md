# excel2dbc-py
**简介**

Inceptio自用DF excel格式转dbc，只针对DF excel模板，生成对应J1939 CAN DBC文件

**使用说明**

- 将待转换的.xlsx文件与.exe放在同一目录下，注意：仅可放一个excel文件
- 点击.exe，即可自动生成.dbc

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