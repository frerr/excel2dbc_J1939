# excel2dbc-py
**简介**

Inceptio自用DF excel格式转dbc，只针对DF excel模板，生成对应J1939 CAN DBC文件

**版本说明**

- 版本：0.1.2，20201016

- 更新点：add logging

- 版本：0.1.1，20200923

- 更新点：fixed .dbc encoding to utf-8
- 版本：0.1.0，20200917
- 更新点：initial release

**EXCEL格式要求**

- 模板格式参见template下DF_Example.xlsx

- **重要1**：Signal start_bit 和 length 在EXCEL里需自行保证分配正确，否则tool生成的dbc，vector_tool会报错

- **重要2**：Signal name和Messege name最大32字符，否则取前32个字符

- **重要3**：Signal value定义，编写格式强制要求如下：

  0x0:Trailer ABS not fully operational
  0x1:Trailer ABS fully operational
  0x2:Error
  0x3:Not available

  其中“：”可替代成“=”，空格或换行符不敏感。禁止出现“0x1-0xF”这种描述方式