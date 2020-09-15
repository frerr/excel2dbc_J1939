# coding=utf-8
#!/usr/bin/python

# Copyright (c) 2020-2030 Wu Bing
# All rights reserved.
#
# Redistribution and use in source and binary forms,
# with or without modification, are permitted provided
# that the following conditions are met:
#
# * Redistributions of source code must retain
#   the above copyright notice, this list of conditions
#   and the following disclaimer.
# * Redistributions in binary form must reproduce
#   the above copyright notice, this list of conditions
#   and the following disclaimer in the documentation
#   and/or other materials provided with the distribution.
# * Neither the name of the author nor the names
#   of its contributors may be used to endorse
#   or promote products derived from this software
#   without specific prior written permission.
#
# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
# "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING,
# BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY
# AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED.
# IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE
# FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY,
# OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
# PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA,
# OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED
# AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT,
# STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
# ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE,
# EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

'''DFdbc excel to DBCfile convertor libraries.'''

def excel2dbc(fin,sheet_name):
    #################################################
    import xlrd
    import re
    import glob

    #################################################
    fout = 'Defalt_out.dbc'

    #################################################
    #                    Read DF-EXCEL              #
    #################################################
    try:
        data = xlrd.open_workbook(fin)

        # 通过index获得工作表info
        table = data.sheet_by_name(sheet_name)

        fout = sheet_name + "_DBC_Out.dbc"

        print(sheet_name+" 总行数：" + str(table.nrows))
        print(sheet_name+ "总列数：" + str(table.ncols))

        # 获取整行的值 和整列的值，返回的结果为数组
        # 整行值：table.row_values(start,end)
        # 整列值：table.col_values(start,end)
        # 参数 start 为从第几个开始打印，
        # end为打印到那个位置结束，默认为none
        # print("整行值：" + str(table.row_values(0)))
        # print("整列值：" + str(table.col_values(9)))

        # rowvalue = table.row_values(4)
        # print(rowvalue)
    except:
        print("Open "+sheet_name +" Error!")
    else:
        print("Open "+sheet_name +" Successfully!")

    #################################################
    #                    Write dbc                  #
    #################################################
    try:
       fdbc = open("dbc"+"/"+fout,"w+")
    except:
        print("Create "+fout +" Error!")
    else:
        print("Create "+fout +" Successfully!")

    ############VERSION##############################
    newContext = "VERSION \"\"\n\n\n"
    fdbc.write(newContext)

    ############NS_ #################################
    newContext = "NS_ : \n"
    fdbc.write(newContext)
    newContext ="        NS_DESC_\n\
            CM_\n\
            BA_DEF_\n\
            BA_\n\
            VAL_\n\
            CAT_DEF_\n\
            CAT_\n\
            FILTER\n\
            BA_DEF_DEF_\n\
            EV_DATA_\n\
            ENVVAR_DATA_\n\
            SGTYPE_\n\
            SGTYPE_VAL_\n\
            BA_DEF_SGTYPE_\n\
            BA_SGTYPE_\n\
            SIG_TYPE_REF_\n\
            VAL_TABLE_\n\
            SIG_GROUP_\n\
            SIG_VALTYPE_\n\
            SIGTYPE_VALTYPE_\n\
            BO_TX_BU_\n\
            BA_DEF_REL_\n\
            BA_REL_\n\
            BA_DEF_DEF_REL_\n\
            BU_SG_REL_\n\
            BU_EV_REL_\n\
            BU_BO_REL_\n\
            SG_MUL_VAL_\n"
    fdbc.write(newContext)

    newContext = "\n"
    fdbc.write(newContext)

    #################BS_BU_网络节点#########################
    try:
        newContext = "BS_:\n"
        fdbc.write(newContext)
        newContext="\n"
        fdbc.write(newContext)
        newContext = "BU_: "
        fdbc.write(newContext)
        Node = table.row_values(0)[30:]
        for Nodename in Node:
            newContext = Nodename+' '
            fdbc.write(newContext)
        newContext="\n\n"
        fdbc.write(newContext)
    except:
        print("Read Node Name Error!")
    else:
        print("Read Node Name Successfully!")
    ################BO_报文###########################
    
    print("print noRow "+str(table.nrows))
    noRow = 2
    spaStr = " "
    chID = "" #MessageID
    Message_name = " "
    Signal_start_bit = 0
    Signal_length = 0
    DLC = 8
    Byte_Order = '1' # 0为intel格式，1为motorola
    Value_type = '+' # + 无符号数， - 有符号数
    Factor = 1.00
    Offset = 0.00
    Real_min_value = 0.00
    Real_max_value = 0.00
    Unit = ""

    while noRow < table.nrows: #每一行遍历

        noRowData = table.row_values(noRow) #每一行的值

        try: 
            if noRowData[2] != "":
                #获取收发节点关系
                Send = ""
                Rev = ""
                for index, i in enumerate(table.row_values(noRow)[30:]):
                    if i == 'S'or i == 's':
                        Send = table.row_values(0)[30+index]
                    elif i == 'r'or i == 'R':
                        Rev += table.row_values(0)[30+index] + ','
                Rev = Rev[:-1]
        except:
            print(noRowData[0]+" Node Rev/Send Relationship Error!")

        try:
            chID = noRowData[2] #第三列为MessageID
            intID = int(chID,16)+int('0x80000000',16) #？？？？？没懂为啥有一位置1了，但dbc文件里实际置了1
            Message_name = noRowData[0]
            DLC = int(noRowData[6])
            #BO_ Message Definition
            newContext="\n"
            fdbc.write(newContext)
            newContext = "BO_ " + str(intID) + spaStr + Message_name + ":" + spaStr + str(DLC) +spaStr +Send+ "\n"
            print("newContext"+newContext)
            fdbc.write(newContext)
        except:
            print(Message_name+"Get Message Info Error!")

        try:
            if noRowData[7] != "":
                Signal_name = noRowData[7]
                Signal_start_bit = int(noRowData[10])
                Signal_length = int(noRowData[13])
                Factor = float(noRowData[15])
                Offset = float(noRowData[16])
                Real_min_value = float(noRowData[17])
                Real_max_value = float(noRowData[18])
                Unit = noRowData[25]
                #SG_ Create
                newContext = spaStr +"SG_" + spaStr + Signal_name + spaStr +":"+ spaStr + str(Signal_start_bit) +"|"+\
                        str(Signal_length) +"@"+ str(Byte_Order) + str(Value_type) + spaStr +"("+\
                        str(Factor) +","+str(Offset) +")"+spaStr+"["+str(Real_min_value) +"|"+str(Real_max_value) +"]"+ \
                        spaStr + spaStr +"\""+Unit+"\"" +spaStr + Rev+ "\n"
                print("newContext" + newContext)
                fdbc.write(newContext)
        except:
            print(" Get Signal Info Error!")
        noRow+=1

    newContext="\n"
    fdbc.write(newContext)

    #########################CM_#######################
    ##           Signal Description                  ##
    ###################################################
    newContext = "CM_ \" \";\n"
    fdbc.write(newContext)

    noRow = 2
    spaStr = " "
    chID = ""
    Signal_name = ""
    Signal_detail = ""

    while noRow < table.nrows: #每一行遍历
        noRowData = table.row_values(noRow) #每一行的值

        if noRowData[2] != "":
            chID = noRowData[2] #第三列为MessageID
            intID = int(chID,16)+int('0x80000000',16) #？？？？？没懂为啥有一位置1了，但dbc文件里实际置了1
    
        #CM_SG_ Create
        if noRowData[7] != "":
            try:
                Signal_name = noRowData[7]
                Signal_detail = noRowData[8]
                newContext ="CM_ "+"SG_ "+str(intID)+spaStr+Signal_name+spaStr+"\""+Signal_detail+"\";"+"\n"
                print("newContext" + newContext)
                fdbc.write(newContext)
            except:
                print(Signal_name+" Get Signal Description Error!")
        noRow+=1

    newContext="\n"
    fdbc.write(newContext)

    #################BA_DEF_BO/SG/BU_##################
    newContext = "BA_DEF_ BO_  \"NmMessage\" ENUM \"No\",\"Yes\";\n\
    BA_DEF_ BO_  \"DiagState\" ENUM  \"No\",\"Yes\";\n\
    BA_DEF_ BO_  \"DiagRequest\" ENUM  \"No\",\"Yes\";\n\
    BA_DEF_ BO_  \"DiagResponse\" ENUM  \"No\",\"Yes\";\n\
    BA_DEF_ BO_  \"GenMsgSendType\" ENUM  \"cyclic\",\"Event\",\"IfActive\",\"OnRequest\",\"CA\",\"CE\",\"NoMsgSendType\";\n\
    BA_DEF_ BO_  \"GenMsgCycleTime\" INT 0 0;\n\
    BA_DEF_ SG_  \"GenSigSendType\" ENUM  \"cyclic\",\"OnChange\",\"OnWrite\",\"IfActive\",\"OnChangeWithRepetition\",\"OnWriteWithRepetition\",\"IfActiveWithRepetition\",\"NoSigSendType\",\"OnChangeAndIfActive\",\"OnChangeAndIfActiveWithRepetition\",\"CA\", \"CE\",\"Event\";\n\
    BA_DEF_ SG_  \"GenSigStartValue\" INT 0 0;\n\
    BA_DEF_ SG_  \"GenSigInactiveValue\" INT 0 0;\n\
    BA_DEF_ BO_  \"GenMsgCycleTimeFast\" INT 0 0;\n\
    BA_DEF_ BO_  \"GenMsgNrOfRepetition\" INT 0 0;\n\
    BA_DEF_ BO_  \"GenMsgDelayTime\" INT 0 0;\n\
    BA_DEF_  \"DBName\" STRING ;\n\
    BA_DEF_ SG_  \"SPN\" INT 0 0;\n\
    BA_DEF_  \"NMIdType\" ENUM  \"0: standard (11 bit, default)\",\"1: extended (29 bit)\";\n\
    BA_DEF_  \"NmMessageCount\" INT 0 256;\n\
    BA_DEF_ BU_  \"NmNode\" ENUM  \"No\",\"Yes\";\n\
    BA_DEF_  \"NmBaseAddress\" HEX 0 536870911;\n\
    BA_DEF_ SG_  \"GenSigEVName\" STRING ;\n\
    BA_DEF_ BO_  \"GenMsgStartDelayTime\" INT 0 10000;\n\
    BA_DEF_ SG_  \"GenSigILSupport\" ENUM  \"Yes\",\"No\";\n\
    BA_DEF_ BO_  \"GenMsgILSupport\" ENUM  \"Yes\",\"No\";\n\
    BA_DEF_ BO_  \"GenMsgRequestable\" INT 0 1;\n\
    BA_DEF_ BO_  \"VFrameFormat\" ENUM  \"StandardCAN\",\"ExtendedCAN\",\"reserved\",\"J1939PG\";\n\
    BA_DEF_ BU_  \"NodeLayerModules\" STRING ;\n\
    BA_DEF_ BU_  \"NmStationAddress\" INT 0 255;\n\
    BA_DEF_ BU_  \"NmJ1939AAC\" INT 0 1;\n\
    BA_DEF_ BU_  \"NmJ1939IndustryGroup\" INT 0 7;\n\
    BA_DEF_ BU_  \"NmJ1939System\" INT 0 127;\n\
    BA_DEF_ BU_  \"NmJ1939SystemInstance\" INT 0 15;\n\
    BA_DEF_ BU_  \"NmJ1939Function\" INT 0 255;\n\
    BA_DEF_ BU_  \"NmJ1939FunctionInstance\" INT 0 7;\n\
    BA_DEF_ BU_  \"NmJ1939ECUInstance\" INT 0 3;\n\
    BA_DEF_ BU_  \"NmJ1939ManufacturerCode\" INT 0 2047;\n\
    BA_DEF_ BU_  \"NmJ1939IdentityNumber\" INT 0 2097151;\n\
    BA_DEF_ BU_  \"ECU\" STRING ;\n\
    BA_DEF_  \"DatabaseVersion\" STRING ;\n\
    BA_DEF_  \"BusType\" STRING ;\n\
    BA_DEF_  \"ProtocolType\" STRING ;\n"
    fdbc.write(newContext)

    newContext="\n"
    fdbc.write(newContext)

    ##################BA_DEF_DEF1#####################
    newContext="BA_DEF_DEF_  \"NmMessage\" \"No\";\n\
    BA_DEF_DEF_  \"DiagState\" \"No\";\n\
    BA_DEF_DEF_  \"DiagRequest\" \"No\";\n\
    BA_DEF_DEF_  \"DiagResponse\" \"No\";\n\
    BA_DEF_DEF_  \"GenMsgSendType\" \"cyclic\";\n\
    BA_DEF_DEF_  \"GenMsgCycleTime\" 0;\n\
    BA_DEF_DEF_  \"GenSigSendType\" \"cyclic\";\n\
    BA_DEF_DEF_  \"GenSigStartValue\" 0;\n"
    fdbc.write(newContext)

    newContext="\n"
    fdbc.write(newContext)

    ###################BA_DEF_DEF2####################
    newContext="BA_DEF_DEF_  \"GenSigInactiveValue\" 0;\n\
    BA_DEF_DEF_  \"GenMsgCycleTimeFast\" 0;\n\
    BA_DEF_DEF_  \"GenMsgNrOfRepetition\" 0;\n\
    BA_DEF_DEF_  \"GenMsgDelayTime\" 0;\n\
    BA_DEF_DEF_  \"DBName\" \"\";\n\
    BA_DEF_DEF_  \"SPN\" 0;\n\
    BA_DEF_DEF_  \"NMIdType\" \"0: standard (11 bit, default)\";\n\
    BA_DEF_DEF_  \"NmMessageCount\" 128;\n\
    BA_DEF_DEF_  \"NmNode\" \"No\";\n\
    BA_DEF_DEF_  \"NmBaseAddress\" 0;\n\
    BA_DEF_DEF_  \"GenSigEVName\" \"Env@Nodename_@Signame\";\n\
    BA_DEF_DEF_  \"GenMsgStartDelayTime\" 0;\n\
    BA_DEF_DEF_  \"GenSigILSupport\" \"Yes\";\n\
    BA_DEF_DEF_  \"GenMsgILSupport\" \"Yes\";\n\
    BA_DEF_DEF_  \"GenMsgRequestable\" 1;\n\
    BA_DEF_DEF_  \"VFrameFormat\" \"J1939PG\";\n\
    BA_DEF_DEF_  \"NodeLayerModules\" \"oseknm01.dll,CANoeILNLVector.dll,J1939_IL.dll\";\n\
    BA_DEF_DEF_  \"NmStationAddress\" 254;\n\
    BA_DEF_DEF_  \"NmJ1939AAC\" 0;\n\
    BA_DEF_DEF_  \"NmJ1939IndustryGroup\" 0;\n\
    BA_DEF_DEF_  \"NmJ1939System\" 0;\n\
    BA_DEF_DEF_  \"NmJ1939SystemInstance\" 0;\n\
    BA_DEF_DEF_  \"NmJ1939Function\" 0;\n\
    BA_DEF_DEF_  \"NmJ1939FunctionInstance\" 0;\n\
    BA_DEF_DEF_  \"NmJ1939ECUInstance\" 0;\n\
    BA_DEF_DEF_  \"NmJ1939ManufacturerCode\" 0;\n\
    BA_DEF_DEF_  \"NmJ1939IdentityNumber\" 0;\n\
    BA_DEF_DEF_  \"ECU\" \"\";\n\
    BA_DEF_DEF_  \"DatabaseVersion\" \"\";\n\
    BA_DEF_DEF_  \"BusType\" \"\";\n\
    BA_DEF_DEF_  \"ProtocolType\" \"\";\n"
    fdbc.write(newContext)

    ##################BA_############################
    newContext="BA_ \"DBName\" \"通讯矩阵- 赢彻B1-I-CAN0724\";\n\
    BA_ \"BusType\" \"J1939\";\n"
    fdbc.write(newContext)

    noRow = 2
    chID = "" #MessageID
    Cycle_time = 0
    SPN = 0

    #BA_ Cycle+SPN
    while noRow < table.nrows: #每一行遍历
        noRowData = table.row_values(noRow) #每一行的值

        if noRowData[2] != "":
            try:
                chID = noRowData[2] #第三列为MessageID
                intID = int(chID,16)+int('0x80000000',16) #？？？？？没懂为啥有一位置1了，但dbc文件里实际置了1
                Cycle_time = int(noRowData[5])
                Message_name = noRowData[0]

                newContext ="BA_ " +"\"GenMsgCycleTime\"" +"BO_ "+str(intID) +spaStr +str(Cycle_time) +";\n"
                print("newContext"+newContext)
                fdbc.write(newContext)
            except:
                print(Message_name+"Cycle Time Error!")

        try:
            if noRowData[7] != "" and int(noRowData[11]) != 0:
                SPN = int(noRowData[11])
                Signal_name = noRowData[7]  
                newContext = "BA_ "+"\"SPN\""+ spaStr + "SG_ "+ str(intID) + spaStr + Signal_name + spaStr + str(SPN)+";\n"
                print("newContext" + newContext)
                fdbc.write(newContext)
        except:
            print(noRowData[7]+"SPN Error!")
        noRow+=1

    newContext="\n"
    fdbc.write(newContext)

    ##############VAL_ value description##############
    noRow = 2
    spaStr = " "
    chID = ""
    Signal_name = ""
    val_str = ""

    while noRow < table.nrows: #每一行遍历
        noRowData = table.row_values(noRow) #每一行的值

        if noRowData[2] != "":
            chID = noRowData[2] #第三列为MessageID
            intID = int(chID,16)+int('0x80000000',16) #？？？？？没懂为啥有一位置1了，但dbc文件里实际置了1
    
        #VAL_ Create
        try:
            if noRowData[7] != "" and noRowData[26] != "":
                Signal_name = noRowData[7]
                line = noRowData[26]
                line = re.sub(r' ','',line)
                list1 = re.split(r'[:,=,\n]',line)

                val_str = ""
                i = 0
                while i < len(list1):
                    if i%2 == 0:
                        val_str += (str(int(list1[i],16))+' ')
                    else:
                        val_str += ("\"" +str(list1[i]) +"\"" +' ')
                    i +=1

                newContext ="VAL_ "+str(intID)+spaStr+Signal_name+spaStr+val_str+";\n"
                print("newContext" + newContext)
                fdbc.write(newContext)
        except:
            print(Signal_name+"Value Definition Error!")
        noRow+=1

    newContext="\n"
    fdbc.write(newContext)

    return fout