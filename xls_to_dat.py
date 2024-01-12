# coding=utf-8
import os
import time
import logging
import xlwings
import unicodedata
from data_set import *
#检查是否有中文
def is_chinese(char):
    if len(char)==0:
        return False
    for i in char:
        if '\u4e00' <= i <= '\u9fff':
            return True
    return False
#检查是否是数字
def is_number(char):
    try:
        float(char)
        return True
    except ValueError:
        pass
    try:
        unicodedata.numeric(char)
        return True
    except (TypeError, ValueError):
        pass
    return False
#检查是否是整数
def is_integer_number(char):
    try:
        int(char)
        return True
    except ValueError:
        pass
    for i in char.split('.')[-1]:
        if i != '0':
            return False
    return True
#检查是否是空白
def is_blank(char):
    if len(char)==0:
        return True
    for i in char:
        if i != ' ':
            return False
    return True
#检查是否整个列表全为空
def is_all_blank(list):
    for i in list:
        if is_blank(i):
            pass
        else:
            return  False
    return True
#检查是否大于一
def is_bigger_one(char):
    if abs(float(char)) >= 1:
        return True
    else:
        return False
#返回整数部分
def get_integer_number(char):
    return int(char.split('.')[0])
#网路检查
def net_check():
    line_list_copy=dat_line_list
    net = line_list_copy[0]
    for i in range(0,len(dat_line_list)):
        char = dat_line_list[i]
        if char[0] in net and char[1] not in net:
            net.append(char[1])
        elif char[1] in net and char[0] not in net:
            net.append(char[0])
    for i in range(0,len(dat_line_list)):
        char = dat_line_list[-i]
        if char[0] in net and char[1] not in net:
            net.append(char[1])
        elif char[1] in net and char[0] not in net:
            net.append(char[0])
    none_net = []
    error_flag = 1
    for i in dat_bus_list:
        if i not in net:
            none_net.append(i)
            error_flag = 0
    if error_flag == 0:
        message=f"已在网络中的节点:"
        for i in net:
            message=message+f"{i},"
        crash_list.append(message)
        message = f"不在网络中的节点:"
        for i in none_net:
            message = message + f"{i},"
        message = message + ""
        crash_list.append(message)
#获取excel的控制卡数据
def get_control_data(sheet_c):
    crash_list.append(f"控制卡的数据：")
    nrows = sheet_c.used_range.last_cell.row#行数
    ncols = sheet_c.used_range.last_cell.column#列数
    control_dict={}
    control_write_data=[]
    contronl_values_list=[]
    contronl_values_list_copy = sheet_c.range('G2:G31').value
    for i in contronl_values_list_copy:
        if i == None:
            contronl_values_list.append('')
        else:
            contronl_values_list.append(str(i))
    for i in range(0,30):
        if i+2 in dat_control_is_chinese_list:
            pass
        elif is_chinese(contronl_values_list[i]):
            crash_list.append(f"<值>所在的第{i+2}行出现了中文！")
        control_dict[dat_control_keys_list[i]] = contronl_values_list[i]
    # 检查潮流方式名是否超过限制与为空
    if len(contronl_values_list[0]) > 10:
        crash_list.append(f"潮流方式名长度超过了限制！")
    elif len(contronl_values_list[0]) == 0:
        crash_list.append(f"潮流方式名不能为空！")
    # 检查工程名是否超过限制与为空
    if len(contronl_values_list[1]) > 20:
        crash_list.append(f"工程名字长度超过了限制！")
    elif len(contronl_values_list[1]) == 0:
        crash_list.append(f"工程名不能为空！")
    # 潮流开始
    CASEID = control_dict['CASEID']
    PROJECT = control_dict['PROJECT']
    control_write_data.append(f"(POWERFLOW,CASEID={CASEID},PROJECT={PROJECT})")
    # 潮流结果输出二进制文件
    NEW_BASE = control_dict['NEW_BASE']
    if is_blank(NEW_BASE):
        pass
    else:
        control_write_data.append(f"/NEW_BASE,FILE={NEW_BASE}.BSE\ ")
    # 指定老库文件
    OLD_BASE = control_dict['OLD_BASE']
    if is_blank(OLD_BASE):
        pass
    else:
        control_write_data.append(f"/OLD_BASE,FILE={OLD_BASE}.BSE\ ")
    # 指定潮流图用的数据文件
    PF_MAP = control_dict['PF_MAP']
    if is_blank(PF_MAP):
        pass
    else:
        control_write_data.append(f"/PF_MAP,FILE = {PF_MAP}.MAP\ ")
    # 系统基准容量
    MVA_BASE = control_dict['MVA_BASE']
    if is_blank(MVA_BASE):
        control_write_data.append(f"/MVA_BASE=100\ ")
    elif is_number(MVA_BASE):
        if is_integer_number(MVA_BASE):
            n = get_integer_number(MVA_BASE)
            control_write_data.append(f"/MVA_BASE={n}\ ")
        else:
            control_write_data.append(f"/MVA_BASE={MVA_BASE}\ ")
    else:
        crash_list.append(f"系统基准容量的值不是数字!")
    # 区域联络线功率控制选择
    AI_CONTROL = control_dict['AI_CONTROL']
    if is_blank(AI_CONTROL):
        pass
    elif AI_CONTROL != 'CON':
        control_write_data.append(f"/AI_CONTROL={AI_CONTROL}\ ")
    # 带负荷调压变压器控制选择
    LTC = control_dict['LTC']
    if is_blank(LTC):
        pass
    elif LTC != 'ON':
        control_write_data.append(f"/LTC={LTC}\ ")
    # 计算方法和迭代次数选择
    SOL_ITER = "/SOL_ITER"
    flag1, flag2, flag3 = 1, 1, 1
    # PQ分解法
    if flag1 == 1:
        DECOUPLED = control_dict['DECOUPLED']
        if is_blank(DECOUPLED):
            flag1 = 0
            SOL_ITER = SOL_ITER + f",DECOUPLED=2"
        elif is_integer_number(DECOUPLED):
            SOL_ITER = SOL_ITER + f",DECOUPLED={get_integer_number(DECOUPLED)}"
            if get_integer_number(DECOUPLED) == 2:
                flag = 0
        else:
            crash_list.append(f"PQ分解法次数填入的值不是正整数!")
    # 改进的牛顿—拉夫逊算法
    if flag2 == 1:
        CURRENT = control_dict['CURRENT']
        if is_blank(CURRENT):
            flag2 = 0
            SOL_ITER = SOL_ITER + f",CURRENT=0"
        elif is_integer_number(CURRENT):
            SOL_ITER = SOL_ITER + f",CURRENT={get_integer_number(CURRENT)}"
            if get_integer_number(CURRENT) == 0:
                flag2 = 0
        else:
            crash_list.append(f"改进的牛顿—拉夫逊算法次数填入的值不是正整数!")
    # 牛顿—拉夫逊算法
    if flag3 == 1:
        NEWTON = control_dict['NEWTON']
        if is_blank(NEWTON):
            flag3 = 0
            SOL_ITER = SOL_ITER + f",NEWTON=30"
        elif is_integer_number(NEWTON):
            SOL_ITER = SOL_ITER + f",NEWTON={get_integer_number(NEWTON)}"
            if get_integer_number(NEWTON) == 30:
                flag3 = 0
        else:
            crash_list.append(f"牛顿—拉夫逊算法次数填入的值不是正整数!")
    # 全为缺省值时才不填写
    if [flag1, flag2, flag3] == [0, 0, 0]:
        pass
    else:
        SOL_ITER = SOL_ITER + "\ "
        control_write_data.append(SOL_ITER)
    # 计算收敛的误差值(标么值)
    TOLERANCE = "/TOLERANCE"
    # 判定是否为缺省值标签
    flag1, flag2, flag3, flag4, flag5 = 1, 1, 1, 1, 1
    # BUSV
    if flag1 == 1:
        BUSV = control_dict['BUSV']
        if is_blank(BUSV):
            flag1 = 0
            TOLERANCE = TOLERANCE + f",BUSV=0.005"
        elif is_number(BUSV):
            if float(BUSV) > 0:
                TOLERANCE = TOLERANCE + f",BUSV={BUSV}"
                if float(BUSV) == 0.005:
                    flag1 = 0
            else:
                crash_list.append(f"计算收敛的误差值(标么值)中BUSV的值不是正数！")
        else:
            crash_list.append(f"计算收敛的误差值(标么值)中BUSV的值不是数字！")
    # AIPOWER
    if flag2 == 1:
        AIPOWER = control_dict['AIPOWER']
        if is_blank(AIPOWER):
            flag2 = 0
            TOLERANCE = TOLERANCE + f",AIPOWER=0.005"
        elif is_number(AIPOWER):
            if float(AIPOWER) > 0:
                TOLERANCE = TOLERANCE + f",AIPOWER={AIPOWER}"
                if float(AIPOWER) == 0.005:
                    flag2 = 0
            else:
                crash_list.append(f"计算收敛的误差值(标么值)中AIPOWER的值不是正数！")
        else:
            crash_list.append(f"计算收敛的误差值(标么值)中AIPOWER的值不是数字！")
    # TX
    if flag3 == 1:
        TX = control_dict['TX']
        if is_blank(TX):
            flag3 = 0
            TOLERANCE = TOLERANCE + f",TX=0.005"
        elif is_number(TX):
            if float(TX) > 0:
                TOLERANCE = TOLERANCE + f",TX={TX}"
                if float(TX) == 0.005:
                    flag3 = 0
            else:
                crash_list.append(f"计算收敛的误差值(标么值)中TX的值不是正数！")
        else:
            crash_list.append(f"计算收敛的误差值(标么值)中TX的值不是数字！")
    # Q
    if flag4 == 1:
        Q = control_dict['Q']
        if is_blank(Q):
            flag4 = 0
            TOLERANCE = TOLERANCE + f",Q=0.005"
        elif is_number(Q):
            if float(Q) > 0:
                TOLERANCE = TOLERANCE + f",Q={Q}"
                if float(Q) == 0.005:
                    flag4 = 0
            else:
                crash_list.append(f"计算收敛的误差值(标么值)中Q的值不是正数！")
        else:
            crash_list.append(f"计算收敛的误差值(标么值)中Q的值不是数字！")
    # OPCUT
    if flag5 == 1:
        OPCUT = control_dict['OPCUT']
        if is_blank(OPCUT):
            flag5 = 0
            TOLERANCE = TOLERANCE + f",OPCUT=0.0001"
        elif is_number(OPCUT):
            if float(OPCUT) > 0:
                TOLERANCE = TOLERANCE + f",OPCUT={OPCUT}"
                if float(OPCUT) == 0.0001:
                    flag5 = 0
            else:
                crash_list.append(f"计算收敛的误差值(标么值)中OPCUT的值不是正数！")
        else:
            crash_list.append(f"计算收敛的误差值(标么值)中OPCUT的值不是数字！")
    # 全都为缺省值才不填写
    if [flag1, flag2, flag3, flag4, flag5] == [0, 0, 0, 0, 0]:
        pass
    else:
        TOLERANCE = TOLERANCE + "\ "
        control_write_data.append(TOLERANCE)
    # 输入数据输出选择
    P_INPUT_LIST = "/P_INPUT_LIST"
    # 缺省值判定标签
    P_INPUT_LIST_flag = 1
    flag1, flag2 = 1, 1
    # 模式
    P_INPUT_LIST_MODEL = control_dict['P_INPUT_LIST_MODEL']
    # 分区名
    INPUT_ZONES = control_dict['INPUT_ZONES']
    # 致命出错是否输出原始数据
    ERRORS = control_dict['ERRORS']
    if P_INPUT_LIST_MODEL in ['FULL', 'ZONES=NONE']:
        P_INPUT_LIST = P_INPUT_LIST + f",{P_INPUT_LIST_MODEL}"
    elif P_INPUT_LIST_MODEL == 'ZONES=ALL，FULL':
        P_INPUT_LIST = P_INPUT_LIST + f",ZONES=ALL"
    elif P_INPUT_LIST_MODEL == 'ZONES=分区名':
        P_INPUT_LIST = P_INPUT_LIST + f",ZONES={INPUT_ZONES}"
    else:
        flag1 = 0
        P_INPUT_LIST = P_INPUT_LIST + f",NONE"
    if ERRORS == "LIST":
        P_INPUT_LIST = P_INPUT_LIST + f",ERRORS={ERRORS}"
    else:
        flag2 = 0
        P_INPUT_LIST = P_INPUT_LIST + f",ERRORS=NO_LIST"
    P_INPUT_LIST = P_INPUT_LIST + "\ "
    # 都为缺省值的时候不填写
    control_write_data.append(P_INPUT_LIST)
    # if flag1 == 0 and flag2 == 0:
    #     P_INPUT_LIST_flag = 0
    #     pass
    # else:
    #     control_write_data.append(P_INPUT_LIST)
    # 潮流计算结果输出选择
    P_OUTPUT_LIST = "/P_OUTPUT_LIST"
    # 缺省值判定标签
    P_OUTPUT_LIST_flag = 1
    flag1, flag2 = 1, 1
    # 模式
    P_OUTPUT_LIST_MODEL = control_dict['P_OUTPUT_LIST_MODEL']
    # 分区名
    OUTPUT_ZONES = control_dict['OUTPUT_ZONES']
    # 迭代不收敛时输出的数据
    FAILED_SOL = control_dict['FAILED_SOL']
    if P_OUTPUT_LIST_MODEL in ['FULL', 'ZONES=NONE']:
        P_OUTPUT_LIST = P_OUTPUT_LIST + f",{P_OUTPUT_LIST_MODEL}"
    elif P_OUTPUT_LIST_MODEL == 'ZONES=ALL，FULL':
        P_OUTPUT_LIST = P_OUTPUT_LIST + f",ZONES=ALL"
    elif P_OUTPUT_LIST_MODEL == 'ZONES=分区名':
        P_OUTPUT_LIST = P_OUTPUT_LIST + f",ZONES={OUTPUT_ZONES}"
    else:
        flag1 = 0
        P_OUTPUT_LIST = P_OUTPUT_LIST + f",NONE"
    if FAILED_SOL in ["PARTIAL_LIST", "NO_LIST"]:
        P_OUTPUT_LIST = P_OUTPUT_LIST + f",FAILED_SOL={ERRORS}"
    else:
        flag2 = 0
        P_OUTPUT_LIST = P_OUTPUT_LIST + f",FAILED_SOL=NO_LIST"
    # 都为缺省值的时候不填写
    P_OUTPUT_LIST = P_OUTPUT_LIST + "\ "
    control_write_data.append(P_OUTPUT_LIST)
    # if flag1 == 0 and flag2 == 0:
    #     P_OUTPUT_LIST_flag = 0
    #     pass
    # else:
    #     control_write_data.append(P_OUTPUT_LIST)
    # 输入数据与结果数据的输出顺序选择
    RPT_SORT = control_dict['RPT_SORT']
    if RPT_SORT in ['ZONE', 'AREA']:
        control_write_data.append(f"/RPT_SORT={RPT_SORT}\ ")
    else:
        control_write_data.append(f"/RPT_SORT=BUS\ ")
    # if [P_INPUT_LIST_flag, P_OUTPUT_LIST_flag] == [0, 0]:
    #     pass
    # else:
    #     if RPT_SORT in ['ZONE', 'AREA']:
    #         control_write_data.append(f"/RPT_SORT={RPT_SORT}\ ")
    #     else:
    #         control_write_data.append(f"/RPT_SORT=BUS\ ")
    # 潮流结果分析报告输出选择
    P_ANALYSIS_RPT = "/P_ANALYSIS_RPT"
    # 缺省值判定标签
    flag1, flag2 = 1, 1
    # 输出等级
    P_ANALYSIS_RPT_LEVEL = control_dict['P_ANALYSIS_RPT_LEVEL']
    # 输出范围
    P_ANALYSIS_RPT_AREA = control_dict['P_ANALYSIS_RPT_AREA']
    # 分区名
    P_ANALYSIS_ZONES = control_dict['P_ANALYSIS_ZONES']
    # 所有者名
    P_ANALYSIS_OWNERS = control_dict['P_ANALYSIS_OWNERS']
    if is_blank(P_ANALYSIS_RPT_LEVEL):
        flag1 = 0
        P_ANALYSIS_RPT = P_ANALYSIS_RPT + f",LEVEL=2"
    elif is_integer_number(P_ANALYSIS_RPT_LEVEL):
        P_ANALYSIS_RPT = P_ANALYSIS_RPT + f",LEVEL={get_integer_number(P_ANALYSIS_RPT_LEVEL)}"
        if get_integer_number(P_ANALYSIS_RPT_LEVEL) == 2:
            flag1 = 0
    if P_ANALYSIS_RPT_AREA == 'ZONES=分区名':
        P_ANALYSIS_RPT = P_ANALYSIS_RPT + f",ZONES={P_ANALYSIS_ZONES}"
    elif P_ANALYSIS_RPT_AREA == 'OWNERS=所有者名':
        P_ANALYSIS_RPT = P_ANALYSIS_RPT + f",OWNERS={P_ANALYSIS_OWNERS}"
    else:
        flag2 = 0
        P_ANALYSIS_RPT = P_ANALYSIS_RPT + f",*"
    # 都为缺省值的时候不填写
    P_ANALYSIS_RPT = P_ANALYSIS_RPT + "\ "
    if flag1 == 0 and flag2 == 0:
        pass
    else:
        control_write_data.append(P_ANALYSIS_RPT)
    # 区域功率交换数据输出控制
    AI_LIST = control_dict['AI_LIST']
    if AI_LIST in ['MATRIX', 'TIELINE', 'NONE']:
        control_write_data.append(f"/AI_LIST={AI_LIST}\ ")
    else:
        # FULL与空白
        pass
    # 输出数据显示精度控制
    OUTPUTDEC = control_dict['OUTPUTDEC']
    if is_blank(OUTPUTDEC):
        pass
    elif is_integer_number(OUTPUTDEC):
        if get_integer_number(OUTPUTDEC) != 2:
            control_write_data.append(f"/OUTPUTDEC={get_integer_number(OUTPUTDEC)}\ ")
    # 支路 R/X 检查控制功能
    RX_CHECK = control_dict['RX_CHECK']
    if RX_CHECK == "OFF":
        control_write_data.append(f"/RX_CHECK=OFF\ ")
    else:
        pass
    control_write_data.append("/NETWORK_DATA\ ")
    return control_write_data
#获取excel的节点数据卡数据
def get_bus_data(sheet_b):
    crash_list.append(f"节点数据卡的数据：")
    nrows = sheet_b.used_range.last_cell.row
    ncols = sheet_b.used_range.last_cell.column
    bus_write_data = []
    bus_write_data.append(".B ----- bus -----")
    for i in range(1, nrows):
        line = str()
        data_range = f"A{i+1}:{excel_range[ncols-1]}{i+1}"
        abus_values_list = []
        abus_values_list_copy = sheet_b.range(data_range).value
        for j in abus_values_list_copy:
            if j == None:
                abus_values_list.append('')
            else:
                abus_values_list.append(str(j))
        if is_all_blank(abus_values_list):
            continue
        for j in range(0,len(abus_values_list)):
            char=abus_values_list[j]
            logth = dat_bus_len_format_list[j]
            # 检查是否有汉字
            if is_chinese(char):
                crash_list.append(f"节点数据卡第{i}行第{j}列数据有汉字!")
            # 空白
            elif is_blank(char):
                line = line + char + ' ' * (logth - len(char))
            # 字母
            elif j in dat_bus_no_digit_list:
                # 得到节点列表
                if j == 3:
                    dat_bus_list.append(char)
                if len(char) > logth:
                    crash_list.append(f"节点数据卡第{i}行第{j}列数据长度超过限制!")
                else:
                    line = line + char + ' ' * (logth - len(char))
            else:
                # 数字
                if is_number(char):
                    if j == 14 and float(char) > 2 and float(char) < 1000:
                        line = line + ("0"+char)[0:logth]
                    # 整数
                    elif is_integer_number(char):
                        char = str(get_integer_number(char))
                        if len(char) > logth:
                            crash_list.append(f"节点数据卡第{i}行第{j}列数据长度超过限制!")
                        else:
                            line = line + (char + '.    ')[0:logth]
                    # 其他数字
                    else:
                        # 大于1
                        if is_bigger_one(char):
                            if len(char) > logth:
                                crash_list.append(f"节点数据卡第{i}行第{j}列数据长度超过限制!")
                            else:
                                line = line + char + ' ' * (logth - len(char))
                        # 小于1
                        else:
                            if len(char) > logth + 1:
                                crash_list.append(f"节点数据卡第{i}行第{j}列数据长度超过限制!")
                            else:
                                line = line + (char + ' ' * (logth + 1 - len(char)))[1:-1] + ' '
                else:
                    crash_list.append(f"节点数据卡第{i}行第{j}列数据不是数字!")
        bus_write_data.append(line)
    return bus_write_data
#获取excel的支路数据卡数据
def get_line_data(sheet_l):
    crash_list.append(f"支路数据卡的数据：")
    nrows = sheet_l.used_range.last_cell.row
    ncols = sheet_l.used_range.last_cell.column
    line_write_data = []
    line_write_data.append(".L ----- transmission lines -----")
    for i in range(1, nrows):
        bus1, bus2 = str(), str()
        line = str()
        data_range = f"A{i + 1}:{excel_range[ncols-1]}{i + 1}"
        aline_values_list = []
        aline_values_list_copy = sheet_l.range(data_range).value
        for j in aline_values_list_copy:
            if j == None:
                aline_values_list.append('')
            else:
                aline_values_list.append(str(j))
        if is_all_blank(aline_values_list):
            continue
        for j in range(0, len(aline_values_list)):
            char = aline_values_list[j]
            logth = dat_line_len_format_list[j]
            # 检查是否有汉字
            if is_chinese(char):
                crash_list.append(f"节点数据卡第{i}行第{j}列数据有汉字!")
            # 空白
            elif is_blank(char):
                line = line + char + ' ' * (logth - len(char))
            # 字母
            elif j in dat_line_no_digit_list:
                # 判定节点是否在节点数据卡中有定义
                if j in [3, 6]:
                    if char not in dat_bus_list:
                        crash_list.append(f"支路数据卡第{i}行第{j}列的节点在节点数据卡中没有定义!")
                    if j == 3:
                        bus1 = char
                    else:
                        bus2 = char
                if len(char) > logth:
                    crash_list.append(f"支路数据卡第{i}行第{j}列数据长度超过限制!")
                else:
                    line = line + char + ' ' * (logth - len(char))
            else:
                # 数字
                if is_number(char):
                    # 整数
                    if is_integer_number(char):
                        char = str(get_integer_number(char))
                        if len(char) > logth:
                            crash_list.append(f"支路数据卡第{i}行第{j}列数据长度超过限制!")
                        else:
                            line = line + (char + '.    ')[0:logth]
                    # 其他数字
                    else:
                        # 大于1
                        if is_bigger_one(char):
                            if len(char) > logth:
                                crash_list.append(f"支路数据卡第{i}行第{j}列数据长度超过限制!")
                            else:
                                line = line + char + ' ' * (logth - len(char))
                        # 小于1
                        else:
                            if len(char) > logth + 1:
                                crash_list.append(f"支路数据卡第{i}行第{j}列数据长度超过限制!")
                            else:
                                line = line + (char + ' ' * (logth + 1 - len(char)))[1:-1] + ' '
                else:
                    crash_list.append(f"支路数据卡第{i}行第{j}列数据不是数字!")
        dat_line_list.append([bus1, bus2])
        line_write_data.append(line)
    return line_write_data
#获取excel的变压器数据卡数据
def get_tran_data(sheet_t):
    crash_list.append(f"变压器数据卡的数据：")
    nrows = sheet_t.used_range.last_cell.row
    ncols = sheet_t.used_range.last_cell.column
    tran_write_data = []
    tran_write_data.append(".T ----- transformers -----")
    for i in range(1, nrows):
        bus1,bus2=str(),str()
        line = str()
        data_range = f"A{i + 1}:{excel_range[ncols-1]}{i + 1}"
        atran_values_list = []
        atran_values_list_copy = sheet_t.range(data_range).value
        for j in atran_values_list_copy:
            if j == None:
                atran_values_list.append('')
            else:
                atran_values_list.append(str(j))
        if is_all_blank(atran_values_list):
            continue
        for j in range(0, len(atran_values_list)):
            char = atran_values_list[j]
            logth = dat_tran_len_format_list[j]
            # 检查是否有汉字
            if is_chinese(char):
                crash_list.append(f"变压器数据卡第{i}行第{j}列数据有汉字!")
            # 空白
            elif is_blank(char):
                line = line + char + ' ' * (logth - len(char))
            # 字母
            elif j in dat_tran_no_digit_list:
                if j in [3, 6]:
                    if char not in dat_bus_list:
                        crash_list.append(f"变压器数据卡第{i}行第{j}列的节点在节点数据卡中没有定义!")
                    if j == 3:
                        bus1 = char
                    else:
                        bus2 = char
                if len(char) > logth:
                    crash_list.append(f"变压器数据卡第{i}行第{j}列数据长度超过限制!")
                else:
                    line = line + char + ' ' * (logth - len(char))
            else:
                # 数字
                if is_number(char):
                    # 整数
                    if is_integer_number(char):
                        char = str(get_integer_number(char))
                        if len(char) > logth:
                            crash_list.append(f"变压器数据卡第{i}行第{j}列数据长度超过限制!")
                        else:
                            line = line + (char + '.    ')[0:logth]
                    # 其他数字
                    else:
                        # 大于1
                        if is_bigger_one(char):
                            if len(char) > logth:
                                crash_list.append(f"变压器数据卡第{i}行第{j}列数据长度超过限制!")
                            else:
                                line = line + char + ' ' * (logth - len(char))
                        # 小于1
                        else:
                            if len(char) > logth + 1:
                                crash_list.append(f"变压器数据卡第{i}行第{j}列数据长度超过限制!")
                            else:
                                line = line + (char + ' ' * (logth + 1 - len(char)))[1:-1] + ' '
                else:
                    crash_list.append(f"变压器数据卡第{i}行第{j}列数据不是数字!")
        dat_line_list.append([bus1,bus2])
        tran_write_data.append(line)
    net_check()
    return  tran_write_data
#写入报错文件
def crash_report():
    if os.path.exists("crashreport"):
        pass
    else:
        os.mkdir("crashreport")
    fn = f"crashreport/{time.strftime('%m%d%H%M%S', time.localtime(time.time()))}.log"
    fm = "%(message)s"
    logging.basicConfig(filename=fn, filemode="w", format=fm, level=logging.DEBUG)
    for i in crash_list:
        logging.debug(i)
    return fn
#数据写入文件
def xls_to_dat(input_file,output_file):
    # EXCEL
    app = xlwings.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    read_workbook =  app.books.open(input_file)
    sheet_c = read_workbook.sheets["控制卡"]
    sheet_b = read_workbook.sheets["节点数据卡"]
    sheet_l = read_workbook.sheets["支路数据卡"]
    sheet_t = read_workbook.sheets["变压器数据卡"]
    list=get_control_data(sheet_c)+get_bus_data(sheet_b)+get_line_data(sheet_l)+get_tran_data(sheet_t)
    list.append("(END)")
    if len(crash_list) == 4:
        with open(output_file,"w+") as f:
            for i in list:
                f.write(i+"\n")
        print(f"文件已输出为:{output_file}")
        input("输入任意值退出")
    else:
        report = crash_report()
        print(f"EXCEL文件数据出错！")
        print(f"详情请查看{report}")
    read_workbook.close()
    app.quit()
if __name__ == "__main__":
    input_file = "input/039bpa.xls"
    output_file = "output/039bpa.DAT"
    xls_to_dat(input_file, output_file)


