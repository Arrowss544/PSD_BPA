# coding=utf-8
import os
import xlwings
import unicodedata
from data_set import *
#检查是否是空白
def is_blank(char):
    if char == None:
        return True
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
#数字读取格式
def number_read_format(number):
    try:
        data=float(number)
    except:
        data=0
    return data
#同步电源转换成直驱风电机组
def gen_to_wp(input_file, output_file):
    # 新建Excle 默认设置：程序可见，只打开不新建工作薄，屏幕更新关闭
    # 创建新的workbook（其实就是创建新的excel）
    app = xlwings.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    read_workbook = app.books.open(input_file)
    read_sheet = read_workbook.sheets["同步电源数据"]
    row_count = read_sheet.api.UsedRange.Rows.Count
    read_data_list = []
    write_data_list = []
    for i in range(0,row_count):
        data_range = f"A{i + 2}:G{i + 2}"
        read_data_list.append(read_sheet.range(data_range).value)
    for i in range(0,len(read_data_list)):
        if is_all_blank(read_data_list[i]):
            pass
        else:
            name = str(str(read_data_list[i][1])+8*" ")[0:8]
            basekv = str(str(read_data_list[i][2])+4*" ")[0:4]
            pgen = float(read_data_list[i][4])
            gen_num = str(1.1 * pgen / 5.5)
            if is_integer_number(gen_num):
                gen_num = (str(int(1.1*pgen/5.5))+3*" ")[0:3]
            else:
                gen_num = (str(int(1.1*pgen/5.5)+1)+3*" ")[0:3]
            new_pgen = (str(int(gen_num)*5.5/1.1)+4*" ")[0:5]
            new_pmax = (str(int(gen_num)*5.5)+4*" ")[0:4]
            new_qmax = (str(int(gen_num)*pow((5.78**2-5.5**2),0.5))+4*" ")[0:5]
            write_data_list.append(f"./节点{name} dat节点数据卡数据")
            B_message = f"BQ    {name}{basekv}00                  {new_pmax}{new_pgen}{new_qmax}0    1."
            write_data_list.append(B_message)
            write_data_list.append("." * 100)
            write_data_list.append(f"./节点{name} swi八张风电机组卡数据")
            MY_meaasge = f"MY {name}{basekv}             5.78                              1000.   0.4 5.5{gen_num}"
            write_data_list.append(MY_meaasge)
            MR_meaasge = f"MR {name}{basekv} 200                                     0.34111001180"
            write_data_list.append(MR_meaasge)
            EU_meaasge = f"EU {name}{basekv}  0.020.418 0.01   1.  -1."
            write_data_list.append(EU_meaasge)
            EZ_meaasge = f"EZ {name}{basekv}            0.02  18.   5. 0.43-0.43 0.15      1              "
            write_data_list.append(EZ_meaasge)
            ES_meaasge = f"ES {name}{basekv}            1.6  1.3  1.4010   2.             0.01 0.01   2.  -2."
            write_data_list.append(ES_meaasge)
            EV_meaasge = f"EV {name}{basekv}  0   0.9 0.91  0.1               "
            write_data_list.append(EV_meaasge)
            LP_meaasge = f"LP {name}{basekv}12  -1.    2  -1.     1 49.1               0           0"
            write_data_list.append(LP_meaasge)
            LQ_meaasge = f"LQ {name}{basekv}   00  0.9  5.30                    01000.  0.2        9000."
            write_data_list.append(LQ_meaasge)
            write_data_list.append("."*100)
    with open(output_file, "w+") as f:
        for i in write_data_list:
            f.write(i + "\n")
    read_workbook.close()
    app.quit()
    print(f"文件已输出为:{output_file}")
    input("输入任意值退出")
if __name__ == "__main__":
    input_file = "input/gen_input.xls"
    output_file = "output/wp_output.swi"
    gen_to_wp(input_file, output_file)