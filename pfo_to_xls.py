# coding=utf-8
import os
import xlwings
import unicodedata
from data_set import *
#检查是否是空白
def is_blank(char):
    if len(char)==0:
        return True
    for i in char:
        if i != ' ':
            return False
    return True
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
#去除中文
def remove_chinese(char):
    new_char = str()
    for i in char:
        if is_chinese(i):
            pass
        else:
            new_char = new_char+i
    return new_char
#获取文本开始特征
def get_startwith_list(char):
    startwith_list = []
    for i in range(1,len(char)+1):
        startwith_list.append(char[0:i])
    return startwith_list
#获取文本结束特征
def get_endwith_list(char):
    endwith_list = []
    if "\n" not in char:
        for i in range(1,len(char)+1):
            endwith_list.append(char[-i:])
    else:
        for i in range(2,len(char)+1):
            endwith_list.append(char[-i:-1])
    return endwith_list
#卡片分类
def card_check(card_list,startwith_list):
    for i in card_list:
        if i in startwith_list:
            return True
    return False
#节点数据
def sheet_B_set(sheet_B_data,sheet_S_data,sheet_B):
    sheet_B_data_list = []
    #获取数据并排序
    for i in range(0,len(sheet_B_data)):
        for j in range(0,len(sheet_B_data)):
            if i == int(sheet_B_data[j].split()[0][3:])-1:
                sheet_B_data_list.append(sheet_B_data[j].split())
    for i in range(0,len(sheet_B_data_list)):
        # 节点名称
        position = "A" + str(i + 2)
        sheet_B.range(position).value = sheet_B_data_list[i][0]
        # 节点类型
        position = "B" + str(i + 2)
        sheet_B.range(position).value = sheet_B_data_list[i][-1]
        # 基准电压（kV）
        position = "C" + str(i + 2)
        sheet_B.range(position).value = sheet_B_data_list[i][1]
        # 实际电压
        position = "D" + str(i + 2)
        sheet_B.range(position).value = sheet_B_data_list[i][2][:-3]
        # 电压标幺值
        position = "E" + str(i + 2)
        sheet_B.range(position).value = sheet_B_data_list[i][-2][:-4]
        # 电压角度
        position = "F" + str(i + 2)
        sheet_B.range(position).value = sheet_B_data_list[i][3][:-1]
        # 其他数据
        for j in range(0,len(sheet_B_data_list[i])):
            if "有功负荷" in sheet_B_data_list[i][j]:
                position = "G" + str(i + 2)
                sheet_B.range(position).value = sheet_B_data_list[i][j][:-4]
            elif "无功负荷" in sheet_B_data_list[i][j]:
                position = "H" + str(i + 2)
                sheet_B.range(position).value = sheet_B_data_list[i][j][:-4]
            elif "有功出力" in sheet_B_data_list[i][j]:
                position = "I" + str(i + 2)
                sheet_B.range(position).value = sheet_B_data_list[i][j][:-4]
            elif "无功出力" in sheet_B_data_list[i][j]:
                position = "J" + str(i + 2)
                sheet_B.range(position).value = sheet_B_data_list[i][j][:-4]
        # 节点有功
        position = "K" + str(i + 2)
        sheet_B.range(position).value = sheet_S_data[i][0][:-4]
        # 节点无功
        position = "L" + str(i + 2)
        sheet_B.range(position).value = sheet_S_data[i][1][:-4]
#支路数据
def sheet_L_set(sheet_L_data,sheet_L):
    range_number = 0
    for i in sheet_L_data:
        for j in sheet_L_data[i]:
            if len(j.split()) == 8 and "/ " in j:
                one_col_list = j.split()[0:3]
                n = j.index(j.split()[2])
                s_char = j[n + len(j.split()[2]):n + len(j.split()[2]) + 48]
                for l in range(0, 4):
                    char = remove_chinese(s_char[l * 12:(l + 1) * 12]).lstrip().rstrip()
                    one_col_list.append(char)
            elif len(j.split()) == 7 and "/ " not in j:
                one_col_list = j.split()[0:3]
                n = j.index(j.split()[2])
                s_char = j[n + len(j.split()[2]):n + len(j.split()[2]) + 48]
                for l in range(0, 4):
                    char = remove_chinese(s_char[l * 12:(l + 1) * 12]).lstrip().rstrip()
                    one_col_list.append(char)
            else:
                one_col_list = j.split()[0:7]
            if float(i[3:]) < float(one_col_list[0].rstrip()[3:]):
                range_number = range_number + 1
                for k in range(0,len(one_col_list)):
                    char = remove_chinese(one_col_list[k])
                    position = excel_range[0] + str(range_number + 1)
                    sheet_L.range(position).value = i
                    position = excel_range[k+1] + str(range_number + 1)
                    #分区名
                    if k == 2:
                        sheet_L.range(position).value = char
                    # 数字
                    elif is_number(char):
                        content = number_read_format(char)
                        sheet_L.range(position).value = content
                    # 字母
                    else:
                        content = char.rstrip()
                        if content != "\n":
                            sheet_L.range(position).value = content
#总结数据
def sheet_SM_set(sheet_SM_data,sheet_SM):
    for i in sheet_SM_data:
        if "发电机出力" in i:
            sheet_SM.range("B3").value = i.split()[1]
            sheet_SM.range("C3").value = i.split()[2]
        elif "负荷" in i:
            sheet_SM.range("B4").value = i.split()[1]
            sheet_SM.range("C4").value = i.split()[2]
        elif "节点并联导纳" in i:
            sheet_SM.range("B5").value = i.split()[1]
            sheet_SM.range("C5").value = i.split()[2]
        elif "未安排的电源" in i:
            sheet_SM.range("B6").value = i.split()[1]
            sheet_SM.range("C6").value = i.split()[2]
        elif "小结（注入）" in i:
            sheet_SM.range("B7").value = i.split()[1]
            sheet_SM.range("C7").value = i.split()[2]
        elif "等值并联导纳" in i:
            sheet_SM.range("E3").value = i.split()[1]
            sheet_SM.range("F3").value = i.split()[2]
        elif "线路和变压器损耗" in i:
            sheet_SM.range("E4").value = i.split()[1]
            sheet_SM.range("F4").value = i.split()[2]
        elif "线路充电功率" in i:
            sheet_SM.range("F5").value = i.split()[1]
        elif "直流换流器损耗" in i:
            sheet_SM.range("E6").value = i.split()[1]
            sheet_SM.range("F6").value = i.split()[2]
        elif "小结（损耗）" in i:
            sheet_SM.range("E7").value = i.split()[1]
            sheet_SM.range("F7").value = i.split()[2]
#写入EXCEL文件
def pfo_to_xls(input_file,output_file):
    # 新建Excle 默认设置：程序可见，只打开不新建工作薄，屏幕更新关闭
    # 创建新的workbook（其实就是创建新的excel）
    app = xlwings.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    write_workbook = app.books.open(pfo_to_xls_template_file)
    if os.path.exists(output_file):
        os.remove(output_file)
    else:
        pass
    write_workbook.save(output_file)
    write_workbook = app.books.open(output_file)
    with open(input_file, 'r', encoding='GBK') as file:
        data = file.readlines()
    # 节点数据
    Bus_type_list = ['B','BE','BQ','BS']
    sheet_B_data = []
    sheet_S_data = []
    sheet_B = write_workbook.sheets["节点数据"]
    # 线路数据
    sheet_L_data = {}
    sheet_L = write_workbook.sheets["支路数据"]
    # 总结数据
    sheet_SM_data = []
    sheet_SM = write_workbook.sheets["总结数据"]
    for i in data:
        endwith_list = get_endwith_list(i)
        if card_check(Bus_type_list,endwith_list):
            if "\n" in i:
                sheet_B_data.append(i[:-1])
            else:
                sheet_B_data.append(i)
        elif "整个系统的数据总结" in i:
            for j in range(data.index(i),len(data)):
                sheet_SM_data.append(data[j][:-1])
                if "系统净输出功率" in data[j]:
                    break
    for i in range(0,len(sheet_B_data)):
        for j in sheet_B_data:
            if i == int(j.split()[0][3:])-1:
                n = data.index(j+"\n")
                for k in range(n,len(data)):
                    if "PNET" in data[k] and "QNET" in data[k]:
                        sheet_L_data[j.split()[0]] = data[n+1:k]
                        sheet_S_data.append(data[k].split())
                        break
    #节点数据
    sheet_B_set(sheet_B_data, sheet_S_data,sheet_B)
    #支路数据
    sheet_L_set(sheet_L_data, sheet_L)
    #总结数据
    sheet_SM_set(sheet_SM_data, sheet_SM)
    write_workbook.save(output_file)
    write_workbook.close()
    app.quit()
    print(f"文件已输出为:{output_file}")
    input("输入任意值退出")
if __name__ == "__main__":
    input_file = "input/039bpa.pfo"
    output_file = "output/039bpa_pfo.xls"
    pfo_to_xls(input_file, output_file)