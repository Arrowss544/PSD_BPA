# coding=utf-8
import os
import xlwings
from data_set import *
#数字读取格式
def number_read_format(number):
    try:
        data=float(number)
    except:
        data=' '
    return data
def new_number_read_format(number):
    try:
        data=float(number)
    except:
        data=0
    return data
#检查是否是空白
def is_blank(char):
    if len(char)==0:
        return True
    for i in char:
        if i != ' ':
            return False
    return True
#是否是中文
def is_chinese(char):
    if '\u4e00' <= char <= '\u9fff':
        return True
    else:
        return False
#处理控制卡命令行
def get_control_content(char):
    for i in range(0,len(char)):
        char="".join(char.split(' '))
    if '(' in char and ')' in char:
        x=char.find('(')+1
        y=char.find(')')
        return char[x:y]
    elif '/' in char and '\\' in char:
        x = char.find('/') + 1
        y = char.find('\\')
        return char[x:y]
    elif '>' in char and '<' in char:
        x = char.find('>') + 1
        y = char.find('<')
        return char[x:y]
#获取文本开始特征
def get_startwith_list(char):
    startwith_list = []
    for i in range(1,len(char)+1):
        startwith_list.append(char[0:i])
    return startwith_list
#读取控制卡数据
def sheet_c_set(sheet_c_data,sheet_c):
    for i in sheet_c_data:
        startwith_list=get_startwith_list(i)
        if "POWERFLOW," in startwith_list:
            one_col_list=i.split(',')
            sheet_c.range('G2').value=one_col_list[1][7:]
            sheet_c.range('G3').value=one_col_list[2][8:]
        elif "NEW_BASE," in startwith_list:
            one_col_list = i.split(',')
            sheet_c.range('G4').value = one_col_list[1][5:-4]
        elif "OLD_BASE," in startwith_list:
            one_col_list = i.split(',')
            sheet_c.range('G5').value = one_col_list[1][5:-4]
        elif "PF_MAP," in startwith_list:
            one_col_list = i.split(',')
            sheet_c.range('G6').value = one_col_list[1][5:-4]
        elif "MVA_BASE" in startwith_list:
            sheet_c.range('G7').value = i[9:]
        elif "AI_CONTROL" in startwith_list:
            sheet_c.range('G8').value = i[11:]
        elif "LTC" in startwith_list:
            sheet_c.range('G9').value = i[4:]
        elif "SOL_ITER" in startwith_list:
            one_col_list = i.split(',')
            for j in one_col_list:
                startwith_list=get_startwith_list(j)
                if "DECOUPLED" in startwith_list:
                    sheet_c.range('G10').value = j[10:]
                elif "CURRENT" in startwith_list:
                    sheet_c.range('G11').value = j[8:]
                elif "NEWTON" in startwith_list:
                    sheet_c.range('G12').value = j[7:]
        elif "TOLERANCE" in startwith_list:
            one_col_list = i.split(',')
            for j in one_col_list:
                startwith_list = get_startwith_list(j)
                if "BUSV" in startwith_list:
                    sheet_c.range('G13').value = j[5:]
                elif "AIPOWER" in startwith_list:
                    sheet_c.range('G14').value = j[8:]
                elif "TX" in startwith_list:
                    sheet_c.range('G15').value = j[3:]
                elif "Q" in startwith_list:
                    sheet_c.range('G16').value = j[2:]
                elif "OPCUT" in startwith_list:
                    sheet_c.range('G17').value = j[6:]
        elif "P_INPUT_LIST" in startwith_list:
            one_col_list = i.split(',')
            one_col_list_copy = []
            one_col_list_copy.append(one_col_list[0])
            if len(one_col_list) >3:
                if "ERRORS" not in one_col_list[-1]:
                    one_col_list_copy.append(",".join(one_col_list[1:]))
                else:
                    one_col_list_copy.append(",".join(one_col_list[1:-1]))
                    one_col_list_copy.append(one_col_list[-1])
            for j in one_col_list_copy:
                startwith_list = get_startwith_list(j)
                if "NONE" in startwith_list:
                    sheet_c.range('G18').value = j
                elif "FULL" in startwith_list:
                    sheet_c.range('G18').value = j
                elif "ZONES=ALL" in startwith_list:
                    sheet_c.range('G18').value = "ZONES=ALL，FULL"
                elif "ZONES=FULL" in startwith_list:
                    sheet_c.range('G18').value = "ZONES=ALL，FULL"
                elif "ZONES=NONE" in startwith_list:
                    sheet_c.range('G18').value = j
                elif "ZONES" in startwith_list:
                    sheet_c.range('G18').value = "ZONES=分区名"
                    sheet_c.range('G19').value = j[6:]
                elif "ERRORS" in startwith_list:
                    sheet_c.range('G20').value = j[7:]
        elif "P_OUTPUT_LIST" in startwith_list:
            one_col_list = i.split(',')
            one_col_list_copy = []
            one_col_list_copy.append(one_col_list[0])
            if len(one_col_list) > 3:
                if "FAILED_SOL" not in one_col_list[-1]:
                    one_col_list_copy.append(",".join(one_col_list[1:]))
                else:
                    one_col_list_copy.append(",".join(one_col_list[1:-1]))
                    one_col_list_copy.append(one_col_list[-1])
            for j in one_col_list_copy:
                startwith_list = get_startwith_list(j)
                if "NONE" in startwith_list:
                    sheet_c.range('G21').value = j
                elif "FULL" in startwith_list:
                    sheet_c.range('G21').value = j
                elif "ZONES=ALL" in startwith_list:
                    sheet_c.range('G21').value = "ZONES=ALL，FULL"
                elif "ZONES=FULL" in startwith_list:
                    sheet_c.range('G21').value = "ZONES=ALL，FULL"
                elif "ZONES=NONE" in startwith_list:
                    sheet_c.range('G21').value = j
                elif "ZONES" in startwith_list:
                    #有问题
                    sheet_c.range('G21').value = "ZONES=分区名"
                    sheet_c.range('G22').value = j[6:]
                elif "FAILED_SOL" in startwith_list:
                    sheet_c.range('G23').value = j[11:]
        elif "RPT_SORT" in startwith_list:
            sheet_c.range('G24').value = i[9:]
        elif "P_ANALYSIS_RPT" in startwith_list:
            one_col_list = i.split(',')
            one_col_list_copy = []
            one_col_list_copy.append(one_col_list[0])
            if len(one_col_list) > 3:
                if "LEVEL" not in one_col_list[1]:
                    one_col_list_copy.append(",".join(one_col_list[1:]))
                else:
                    one_col_list_copy.append(",".join(one_col_list[2:]))
                    one_col_list_copy.append(one_col_list[1])
            for j in  one_col_list_copy:
                startwith_list = get_startwith_list(j)
                if "LEVEL" in startwith_list:
                    sheet_c.range('G25').value = j[6:]
                elif "*" in startwith_list:
                    sheet_c.range('G26').value = j
                elif "ZONES" in startwith_list:
                    sheet_c.range('G26').value = "ZONES=分区名"
                    sheet_c.range('G27').value = j[6:]
                elif "OWNERS" in startwith_list:
                    sheet_c.range('G26').value = "OWNERS=所有者名"
                    sheet_c.range('G28').value = j[7:]
        elif "AI_LIST" in startwith_list:
            sheet_c.range('G29').value = i[8:]
        elif "OUTPUTDEC" in startwith_list:
            one_col_list = i.split(',')
            sheet_c.range('G30').value = one_col_list[1][6:]
        elif "CHECK" in startwith_list:
            one_col_list = i.split(',')
            sheet_c.range('G31').value = one_col_list[1][9:]
#读取节点数据卡数据
def sheet_b_set(sheet_b_data,sheet_b,sheet_b_s):
    sheet_b_data_list=[]
    sheet_nb_data_list = []
    sheet_lb_data_list = []
    sheet_gb_data_list = []
    for i in range(0,len(sheet_b_data)):
        one_col_b_data_list=[]
        n=0
        for j in dat_bus_len_format_list:
            one_col_b_data_list.append(sheet_b_data[i][n:(n+int(j))])
            n = n + int(j)
        sheet_b_data_list.append(one_col_b_data_list)
    for i in range(0,len(sheet_b_data_list)):
        for j in range(0,len(sheet_b_data_list[i])):
            position = excel_range[j]+str(i+2)
            #字母
            if j in dat_bus_no_digit_list:
                content = sheet_b_data_list[i][j].rstrip()
                if content != "\n":
                    sheet_b.range(position).value=content
            #数字
            else:
                content = number_read_format(sheet_b_data_list[i][j])
                sheet_b.range(position).value=content
        name = sheet_b_data_list[i][3].rstrip()
        load_p = new_number_read_format(sheet_b_data_list[i][6])
        load_q = new_number_read_format(sheet_b_data_list[i][7])
        gen_p = new_number_read_format(sheet_b_data_list[i][11])
        gen_q = new_number_read_format(sheet_b_data_list[i][12])
        if [load_p,load_q,gen_p,gen_q] == [0,0,0,0]:
            sheet_nb_data_list.append([name,])
        elif [gen_p,gen_q] == [0,0]:
            sheet_lb_data_list.append([name,load_p,load_q])
        else:
            sheet_gb_data_list.append([name,gen_p,gen_q,gen_p-load_p,gen_q-load_q])
    sheet_b_s.range("A3").value = sheet_nb_data_list
    sheet_b_s.range("B3").value = sheet_lb_data_list
    sheet_b_s.range("E3").value = sheet_gb_data_list
#读取支路数据卡数据
def sheet_l_set(sheet_l_data,sheet_l):
    sheet_l_data_list = []
    for i in range(0, len(sheet_l_data)):
        one_col_l_data_list = []
        n = 0
        for j in dat_line_len_format_list:
            one_col_l_data_list.append(sheet_l_data[i][n:(n + int(j))])
            n = n + int(j)
        sheet_l_data_list.append(one_col_l_data_list)
    for i in range(0, len(sheet_l_data_list)):
        for j in range(0, len(sheet_l_data_list[i])):
            position = excel_range[j] + str(i + 2)
            # 字母
            if j in dat_line_no_digit_list:
                content = sheet_l_data_list[i][j].rstrip()
                if content != "\n":
                    sheet_l.range(position).value=content
            # 数字
            else:
                content = number_read_format(sheet_l_data_list[i][j])
                sheet_l.range(position).value=content
#读取变压器数据卡数据
def sheet_t_set(sheet_t_data,sheet_t):
    sheet_t_data_list = []
    for i in range(0, len(sheet_t_data)):
        one_col_t_data_list = []
        n = 0
        for j in dat_tran_len_format_list:
            one_col_t_data_list.append(sheet_t_data[i][n:(n + int(j))])
            n = n + int(j)
        sheet_t_data_list.append(one_col_t_data_list)
    for i in range(0, len(sheet_t_data_list)):
        for j in range(0, len(sheet_t_data_list[i])):
            position = excel_range[j] + str(i + 2)
            # 字母
            if j in dat_tran_no_digit_list:
                content = sheet_t_data_list[i][j].rstrip()
                if content != "\n":
                    sheet_t.range(position).value=content
            # 数字
            else:
                content = number_read_format(sheet_t_data_list[i][j])
                sheet_t.range(position).value=content
#写入EXCEL文件
def dat_to_xls(input_file,output_file):
    # 新建Excle 默认设置：程序可见，只打开不新建工作薄，屏幕更新关闭
    # 创建新的workbook（其实就是创建新的excel）
    app = xlwings.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    write_workbook = app.books.open(dat_to_xls_template_file)
    if os.path.exists(output_file):
        os.remove(output_file)
    else:
        pass
    write_workbook.save(output_file)
    write_workbook = app.books.open(output_file)
    with open(input_file, 'r', encoding='GBK') as file:
        data = file.readlines()
    # 控制卡数据
    sheet_c_data = []
    # 节点数据卡数据
    sheet_b_data = []
    # 支路数据卡数据
    sheet_l_data = []
    # 变压器数据卡数据
    sheet_t_data = []
    sheet_c = write_workbook.sheets["控制卡"]
    sheet_b = write_workbook.sheets["节点数据卡"]
    sheet_l = write_workbook.sheets["支路数据卡"]
    sheet_t = write_workbook.sheets["变压器数据卡"]
    sheet_b_s = write_workbook.sheets["节点数据统计"]
    for i in data:
        if i[0] not in ['.', '(', '/', '>']:
            if i[0:2] in ['B ', 'BQ', 'BE', 'BS']:
                sheet_b_data.append(i + ' ' * (sum(dat_bus_len_format_list)+1-len(i)))
            elif i[0:2] in ['L ']:
                sheet_l_data.append(i + ' ' * (sum(dat_line_len_format_list)+1-len(i)))
            elif i[0:2] in ['T ']:
                sheet_t_data.append(i + ' ' * (sum(dat_tran_len_format_list)+1-len(i)))
        if i[0] in [ '(', '/', '>']:
            sheet_c_data.append(get_control_content(i))
    sheet_c_set(sheet_c_data,sheet_c)
    sheet_b_set(sheet_b_data,sheet_b,sheet_b_s)
    sheet_l_set(sheet_l_data,sheet_l)
    sheet_t_set(sheet_t_data,sheet_t)
    write_workbook.save()
    write_workbook.close()
    app.quit()
    print(f"文件已输出为:{output_file}")
    input("输入任意值退出")
if __name__ == "__main__":
    input_file = "input/039bpa.DAT"
    output_file = "output/039bpa_DAT.xls"
    dat_to_xls(input_file, output_file)
