# coding=utf-8
import os
import xlwings
import unicodedata
from data_set import *
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
#数字读取格式
def number_read_format(number):
    try:
        data=float(number)
    except:
        data=' '
    return data
#是否是中文
def is_chinese(char):
    if '\u4e00' <= char <= '\u9fff':
        return True
    else:
        return False
#获取文本开始特征
def get_startwith_list(char):
    startwith_list = []
    for i in range(1,len(char)+1):
        startwith_list.append(char[0:i])
    return startwith_list
#卡片分类
def card_check(card_list,startwith_list):
    for i in card_list:
        if i in startwith_list:
            return True
    return False
#模板写入sheet
def template_sheet_set(sheet_data,sheet,len_format_list,is_blank_list):
    for i in range(0, len(sheet_data)):
        n = 0
        range_number = 0
        for j in range(0, len(len_format_list)):
            m = len_format_list[j]
            char = sheet_data[i][n:(n + m)]
            n = n + m
            if j not in is_blank_list:
                position = excel_range[range_number] + str(i + 2)
                range_number = range_number + 1
                # 数字
                if is_number(char):
                    content = number_read_format(char)
                    sheet.range(position).value = content
                # 字母
                else:
                    content = char.rstrip()
                    if content != "\n":
                        sheet.range(position).value = content
# 控制卡数据
# 计算控制CASE卡数据
# 监视曲线控制F0卡数据
# 计算控制继续F1卡数据
# 计算控制FF卡数据
def sheet_C_set(sheet_C_data,sheet_C):
    logth = 3
    for i in range(0, len(sheet_C_data)):
        startwith_list = get_startwith_list(sheet_C_data[i])
        n = 0
        if "CASE" in startwith_list:
            range_number = 0
            for j in range(0, len(swi_CASE_len_format_list)):
                m = swi_CASE_len_format_list[j]
                char = sheet_C_data[i][n:(n + m)]
                n = n + m
                if j not in swi_CASE_is_blank_list:
                    position = excel_range[range_number] + str(3)
                    range_number = range_number+1
                    # 数字
                    if is_number(char):
                        content = number_read_format(char)
                        sheet_C.range(position).value = content
                    # 字母
                    else:
                        content = char.rstrip()
                        if content != "\n":
                            sheet_C.range(position).value = content
        elif "F0" in startwith_list:
            range_number = 0
            for j in range(0, len(swi_F0_len_format_list)):
                m = swi_F0_len_format_list[j]
                char = sheet_C_data[i][n:(n + m)]
                n = n + m
                if j not in swi_F0_is_blank_list:
                    position = excel_range[range_number] + str(3 + 1*logth)
                    range_number = range_number +1
                    # 数字
                    if is_number(char):
                        content = number_read_format(char)
                        sheet_C.range(position).value = content
                    # 字母
                    else:
                        content = char.rstrip()
                        if content != "\n":
                            sheet_C.range(position).value = content
        elif "F1" in startwith_list:
            range_number = 0
            for j in range(0, len(swi_F1_len_format_list)):
                m = swi_F1_len_format_list[j]
                char = sheet_C_data[i][n:(n + m)]
                n = n + m
                if j not in swi_F1_is_blank_list:
                    position = excel_range[range_number] + str(3 + 2*logth)
                    range_number = range_number +1
                    # 数字
                    if is_number(char):
                        content = number_read_format(char)
                        sheet_C.range(position).value = content
                    # 字母
                    else:
                        content = char.rstrip()
                        if content != "\n":
                            sheet_C.range(position).value = content
        elif "FF" in startwith_list:
            range_number = 0
            for j in range(0, len(swi_FF_len_format_list)):
                m = swi_FF_len_format_list[j]
                char = sheet_C_data[i][n:(n + m)]
                n = n + m
                if j not in swi_FF_is_blank_list:
                    position = excel_range[range_number] + str(3 + 3*logth)
                    range_number = range_number +1
                    # 数字
                    if is_number(char):
                        content = number_read_format(char)
                        sheet_C.range(position).value = content
                    # 字母
                    else:
                        content = char.rstrip()
                        if content != "\n":
                            sheet_C.range(position).value = content
# 基本故障模型LS卡数据
def sheet_LS_set(sheet_LS_data,sheet_LS):
    pass
#  风电机组模型MY卡数据
#  风电机组模型下的MR卡——低电压穿越保护模型
#  风电机组模型下的EU卡——网侧变频器有功控制模型
#  风电机组模型下的EZ卡——正常运行状态下无功控制模型
#  风电机组模型下的ES卡——有功无功电流限制模型
#  风电机组模型下的EV卡——低电压高电压状态判断模型
#  风电机组模型下的LP卡——低电压穿越状态下有功控制模型
#  风电机组模型下的LQ卡——低电压穿越状态下无功控制模型
def sheet_WP_set(sheet_WT_data,sheet_WT):
    logth = int(len(sheet_WT_data)/8)+2
    for i in range(0,len(sheet_WT_data)):
        startwith_list = get_startwith_list(sheet_WT_data[i])
        n = 0
        MY_number = 0
        MR_number = 0
        EU_number = 0
        EZ_number = 0
        ES_number = 0
        EV_number = 0
        LP_number = 0
        LQ_number = 0
        if "MY" in startwith_list:
            range_number = 0
            MY_number = MY_number + 1
            for j in range(0,len(swi_MY_len_format_list)):
                m = swi_MY_len_format_list[j]
                char = sheet_WT_data[i][n:(n + m)]
                n = n + m
                if j not in swi_MY_is_blank_list:
                    position = excel_range[range_number] + str(MY_number + 2)
                    range_number = range_number +1
                    # 数字
                    if is_number(char):
                        content = number_read_format(char)
                        sheet_WT.range(position).value = content
                    # 字母
                    else:
                        content = char.rstrip()
                        if content != "\n":
                            sheet_WT.range(position).value = content
        elif "MR" in startwith_list:
            range_number = 0
            MR_number = MR_number + 1
            for j in range(0, len(swi_MR_len_format_list)):
                m = swi_MR_len_format_list[j]
                char = sheet_WT_data[i][n:(n + m)]
                n = n + m
                if j not in swi_MR_is_blank_list:
                    position = excel_range[range_number] + str(MR_number + 2 + 1*logth)
                    range_number = range_number +1
                    # 数字
                    if is_number(char):
                        content = number_read_format(char)
                        sheet_WT.range(position).value = content
                    # 字母
                    else:
                        content = char.rstrip()
                        if content != "\n":
                            sheet_WT.range(position).value = content
        elif "EU" in startwith_list:
            range_number = 0
            EU_number = EU_number + 1
            for j in range(0, len(swi_EU_len_format_list)):
                m = swi_EU_len_format_list[j]
                char = sheet_WT_data[i][n:(n + m)]
                n = n + m
                if j not in swi_EU_is_blank_list:
                    position = excel_range[range_number] + str(EU_number + 2 + 2*logth)
                    range_number = range_number + 1
                    # 数字
                    if is_number(char):
                        content = number_read_format(char)
                        sheet_WT.range(position).value = content
                    # 字母
                    else:
                        content = char.rstrip()
                        if content != "\n":
                            sheet_WT.range(position).value = content
        elif "EZ" in startwith_list:
            range_number = 0
            EZ_number = EZ_number + 1
            for j in range(0, len(swi_EZ_len_format_list)):
                m = swi_EZ_len_format_list[j]
                char = sheet_WT_data[i][n:(n + m)]
                n = n + m
                if j not in swi_EZ_is_blank_list:
                    position = excel_range[range_number] + str(EZ_number + 2 + 3*logth)
                    range_number = range_number +1
                    # 数字
                    if is_number(char):
                        content = number_read_format(char)
                        sheet_WT.range(position).value = content
                    # 字母
                    else:
                        content = char.rstrip()
                        if content != "\n":
                            sheet_WT.range(position).value = content
        elif "ES" in startwith_list:
            range_number = 0
            ES_number = ES_number + 1
            for j in range(0, len(swi_ES_len_format_list)):
                m = swi_ES_len_format_list[j]
                char = sheet_WT_data[i][n:(n + m)]
                n = n + m
                if j not in swi_ES_is_blank_list:
                    position = excel_range[range_number] + str(ES_number + 2 + 4*logth)
                    range_number = range_number +1
                    # 数字
                    if is_number(char):
                        content = number_read_format(char)
                        sheet_WT.range(position).value = content
                    # 字母
                    else:
                        content = char.rstrip()
                        if content != "\n":
                            sheet_WT.range(position).value = content
        elif "EV" in startwith_list:
            range_number = 0
            EV_number = EV_number + 1
            for j in range(0, len(swi_EV_len_format_list)):
                m = swi_EV_len_format_list[j]
                char = sheet_WT_data[i][n:(n + m)]
                n = n + m
                if j not in swi_EV_is_blank_list:
                    position = excel_range[range_number] + str(EV_number + 2 + 5*logth)
                    range_number = range_number +1
                    # 数字
                    if is_number(char):
                        content = number_read_format(char)
                        sheet_WT.range(position).value = content
                    # 字母
                    else:
                        content = char.rstrip()
                        if content != "\n":
                            sheet_WT.range(position).value = content
        elif "LP" in startwith_list:
            range_number = 0
            LP_number = LP_number + 1
            for j in range(0, len(swi_LP_len_format_list)):
                m = swi_LP_len_format_list[j]
                char = sheet_WT_data[i][n:(n + m)]
                n = n + m
                if j not in swi_LP_is_blank_list:
                    position = excel_range[range_number] + str(LP_number + 2 + 6*logth)
                    range_number = range_number +1
                    # 数字
                    if is_number(char):
                        content = number_read_format(char)
                        sheet_WT.range(position).value = content
                    # 字母
                    else:
                        content = char.rstrip()
                        if content != "\n":
                            sheet_WT.range(position).value = content
        elif "LQ" in startwith_list:
            range_number = 0
            LQ_number = LQ_number + 1
            for j in range(0, len(swi_LQ_len_format_list)):
                m = swi_LQ_len_format_list[j]
                char = sheet_WT_data[i][n:(n + m)]
                n = n + m
                if j not in swi_LQ_is_blank_list:
                    position = excel_range[range_number] + str(LQ_number + 2 + 7*logth)
                    range_number = range_number +1
                    # 数字
                    if is_number(char):
                        content = number_read_format(char)
                        sheet_WT.range(position).value = content
                    # 字母
                    else:
                        content = char.rstrip()
                        if content != "\n":
                            sheet_WT.range(position).value = content
# 输出控制卡数据
# 输出主控制MH卡数据
# 母线输出控制BH卡数据
# 发电机输出控制GH卡数据
def sheet_OC_set(sheet_OC_data,sheet_OC):
    logth = 3
    for i in range(0, len(sheet_OC_data)):
        startwith_list = get_startwith_list(sheet_OC_data[i])
        n = 0
        if "MH" in startwith_list:
            range_number = 0
            for j in range(0, len(swi_MH_len_format_list)):
                m = swi_MH_len_format_list[j]
                char = sheet_OC_data[i][n:(n + m)]
                n = n + m
                if j not in swi_MH_is_blank_list:
                    position = excel_range[range_number] + str(3)
                    range_number = range_number + 1
                    # 数字
                    if is_number(char):
                        content = number_read_format(char)
                        sheet_OC.range(position).value = content
                    # 字母
                    else:
                        content = char.rstrip()
                        if content != "\n":
                            sheet_OC.range(position).value = content
        elif "BH" in startwith_list:
            range_number = 0
            for j in range(0, len(swi_BH_len_format_list)):
                m = swi_BH_len_format_list[j]
                char = sheet_OC_data[i][n:(n + m)]
                n = n + m
                if j not in swi_BH_is_blank_list:
                    position = excel_range[range_number] + str(3 + 1 * logth)
                    range_number = range_number + 1
                    # 数字
                    if is_number(char):
                        content = number_read_format(char)
                        sheet_OC.range(position).value = content
                    # 字母
                    else:
                        content = char.rstrip()
                        if content != "\n":
                            sheet_OC.range(position).value = content
        elif "GH" in startwith_list:
            range_number = 0
            for j in range(0, len(swi_GH_len_format_list)):
                m = swi_GH_len_format_list[j]
                char = sheet_OC_data[i][n:(n + m)]
                n = n + m
                if j not in swi_GH_is_blank_list:
                    position = excel_range[range_number] + str(3 + 2 * logth)
                    range_number = range_number + 1
                    # 数字
                    if is_number(char):
                        content = number_read_format(char)
                        sheet_OC.range(position).value = content
                    # 字母
                    else:
                        content = char.rstrip()
                        if content != "\n":
                            sheet_OC.range(position).value = content
#写入EXCEL文件
def swi_to_xls(input_file,output_file):
    # 新建Excle 默认设置：程序可见，只打开不新建工作薄，屏幕更新关闭
    # 创建新的workbook（其实就是创建新的excel）
    app = xlwings.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    write_workbook = app.books.open(swi_to_xls_template_file)
    if os.path.exists(output_file):
        os.remove(output_file)
    else:
        pass
    write_workbook.save(output_file)
    write_workbook = app.books.open(output_file)
    with open(input_file, 'r', encoding='GBK') as file:
        data = file.readlines()
    # 第一个部分填写所有模型、故障操作参数。
    # 计算控制CASE卡数据
    # 计算控制继续F1卡数据
    # 监视曲线控制F0卡数据
    # 计算控制FF卡数据
    swi_control_list = ["CASE","F1","F0","FF"]
    sheet_C_data = []
    sheet_C = write_workbook.sheets["计算控制卡"]
    # 基本故障模型LS卡数据
    sheet_LS_data = []
    #sheet_LS = write_workbook.sheets["基本故障模型LS卡",after="计算控制卡")
    #  风电机组模型MY卡数据
    #  风电机组模型下的EZ卡——正常运行状态下无功控制模型
    #  风电机组模型下的ES卡——有功无功电流限制模型
    #  风电机组模型下的EV卡——低电压高电压状态判断模型
    #  风电机组模型下的LP卡——低电压穿越状态下有功控制模型
    #  风电机组模型下的LQ卡——低电压穿越状态下无功控制模型
    #  风电机组模型下的EU卡——网侧变频器有功控制模型
    #  风电机组模型下的MR卡——低电压穿越保护模型
    swi_wind_turbine_list = ["MY","EZ","ES","EV","LP","LQ","EU","MR"]
    sheet_WT_data = []
    sheet_WT = write_workbook.sheets["风电机组模型数据"]
    #发电机次暂态参数模型M卡数据
    sheet_M_data = []
    sheet_M = write_workbook.sheets["发电机次暂态参数模型M卡"]
    # 发电机双轴模型MF卡数据
    sheet_MF_data = []
    sheet_MF = write_workbook.sheets["发电机双轴模型MF卡"]
    # 发电机E'恒定模型MC卡数据
    sheet_MC_data = []
    sheet_MC = write_workbook.sheets["发电机E'恒定模型MC卡"]
    # 静态负荷模型LB卡数据
    sheet_LB_data = []
    sheet_LB = write_workbook.sheets["静态负荷模型LB卡"]
    # 第二个部分填写输出数据。
    # 输出控制卡数据
    # 输出主控制MH卡数据
    # 母线输出控制BH卡数据
    # 发电机输出控制GH卡数据
    swi_output_control_list = ['MH','BH','GH']
    sheet_OC_date = []
    sheet_OC = write_workbook.sheets["输出控制卡"]
    # 母线输出B卡数据
    sheet_B_data = []
    sheet_B = write_workbook.sheets["母线输出B卡"]
    # 发电机输出G卡数据
    sheet_G_data = []
    sheet_G = write_workbook.sheets["发电机输出G卡"]
    for i in data:
        startwith_list = get_startwith_list(i)
        if card_check(swi_control_list,startwith_list):
            sheet_C_data.append(i+' '*80)
        elif "LS" in startwith_list:
            sheet_LS_data.append(i+' '*80)
        elif card_check(swi_wind_turbine_list,startwith_list):
            sheet_WT_data.append(i+' '*80)
        elif "M " in startwith_list:
            sheet_M_data.append(i+' '*80)
        elif "MF" in startwith_list:
            sheet_MF_data.append(i+' '*80)
        elif "MC" in startwith_list:
            sheet_MC_data.append(i+' '*80)
        elif "LB" in startwith_list:
            sheet_LB_data.append(i+' '*80)
        elif card_check(swi_output_control_list,startwith_list):
            sheet_OC_date.append(i+' '*80)
        elif "B " in startwith_list:
            sheet_B_data.append(i+' '*80)
        elif "G " in startwith_list:
            sheet_G_data.append(i+' '*80)
    # 控制卡数据
    # 计算控制CASE卡数据
    # 监视曲线控制F0卡数据
    # 计算控制继续F1卡数据
    # 计算控制FF卡数据
    sheet_C_set(sheet_C_data, sheet_C)
    # 基本故障模型LS卡数据
    #sheet_LS_set(sheet_LS_data, sheet_LS)
    #  风电机组模型MY卡数据
    #  风电机组模型下的EZ卡——正常运行状态下无功控制模型
    #  风电机组模型下的ES卡——有功无功电流限制模型
    #  风电机组模型下的EV卡——低电压高电压状态判断模型
    #  风电机组模型下的LP卡——低电压穿越状态下有功控制模型
    #  风电机组模型下的LQ卡——低电压穿越状态下无功控制模型
    #  风电机组模型下的EU卡——网侧变频器有功控制模型
    #  风电机组模型下的MR卡——低电压穿越保护模型
    sheet_WP_set(sheet_WT_data, sheet_WT)
    # 读取发电机次暂态参数M卡模型
    template_sheet_set(sheet_M_data,sheet_M,swi_M_len_format_list,swi_M_is_blank_list)
    #sheet_M_set(sheet_M_data, sheet_M)
    # 发电机双轴模型MF卡数据
    template_sheet_set(sheet_MF_data, sheet_MF, swi_MC_or_MF_len_format_list, swi_MC_or_MF_is_blank_list)
    #sheet_MF_set(sheet_MF_data, sheet_MF)
    # 发电机E'恒定模型MC卡数据
    template_sheet_set(sheet_MC_data, sheet_MC, swi_MC_or_MF_len_format_list, swi_MC_or_MF_is_blank_list)
    #sheet_MC_set(sheet_MC_data, sheet_MC)
    # 静态负荷模型LB卡数据
    template_sheet_set(sheet_LB_data, sheet_LB, swi_LB_or_LC_len_format_list, swi_LB_or_LC_is_blank_list)
    # sheet_LB_set(sheet_LB_data, sheet_LB)
    # 输出控制卡数据
    # 输出主控制MH卡数据
    # 母线输出控制BH卡数据
    # 发电机输出控制GH卡数据
    sheet_OC_set(sheet_OC_date, sheet_OC)
    # 母线输出B卡数据
    template_sheet_set(sheet_B_data,sheet_B,swi_B_len_format_list,swi_B_is_blank_list)
    # 发电机输出G卡数据
    template_sheet_set(sheet_G_data,sheet_G,swi_G_len_format_list,swi_G_is_blank_list)
    write_workbook.save(output_file)
    write_workbook.close()
    app.quit()
    print(f"文件已输出为:{output_file}")
    input("输入任意值退出")
if __name__ == "__main__":
    input_file = "input/039bpa.swi"
    output_file = "output/039bpa_swi.xls"
    swi_to_xls(input_file, output_file)