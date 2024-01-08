# coding=utf-8
import os
import xlwings
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
#数字读取格式
def number_read_format(number):
    try:
        data=float(number)
    except:
        data=0
    return data
#标幺值转换
def pu_to_pu(input_file, output_file):
    # 新建Excle 默认设置：程序可见，只打开不新建工作薄，屏幕更新关闭
    # 创建新的workbook（其实就是创建新的excel）
    app = xlwings.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    write_workbook = app.books.open(pu_to_pu_template_file)
    if os.path.exists(output_file):
        os.remove(output_file)
    else:
        pass
    write_workbook.save(output_file)
    write_workbook = app.books.open(output_file)
    read_workbook = app.books.open(input_file)
    write_sheet = write_workbook.sheets["标幺值转换"]
    read_sheet = read_workbook.sheets["标幺值转换"]
    row_count = read_sheet.api.UsedRange.Rows.Count
    basekv = int(read_sheet.range("C1").value)
    basemva = int(read_sheet.range("F1").value)
    read_data_list = []
    wriet_data_list = []
    for i in range(0,row_count):
        data_range = f"A{i + 3}:I{i + 3}"
        read_data_list.append(read_sheet.range(data_range).value)
    for i in range(0,len(read_data_list)):
        if is_all_blank(read_data_list[i]):
            pass
        else:
            one_wriet_data_list = read_data_list[i]
            ratedkv = int(read_data_list[i][3])
            ratedmva = int(read_data_list[i][4])
            rpu = float(number_read_format(read_data_list[i][5]))
            xpu = float(number_read_format(read_data_list[i][6]))
            gpu = float(number_read_format(read_data_list[i][7]))
            bpu = float(number_read_format(read_data_list[i][8]))
            nrpu = (rpu * basemva * (ratedkv ** 2)) / (ratedmva * (basekv ** 2))
            nxpu = (xpu * basemva * (ratedkv ** 2)) / (ratedmva * (basekv ** 2))
            ngpu = (gpu * ratedmva * (basekv ** 2)) / (basemva * (ratedkv ** 2))
            nbpu = (bpu * ratedmva * (basekv ** 2)) / (basemva * (ratedkv ** 2))
            one_wriet_data_list.append(nrpu)
            one_wriet_data_list.append(nxpu)
            one_wriet_data_list.append(ngpu)
            one_wriet_data_list.append(nbpu)
            wriet_data_list.append(one_wriet_data_list)
        write_sheet.range("C1").value = basekv
        write_sheet.range("F1").value = basemva
        for i in range(0,len(wriet_data_list)):
            data_range = f"A{i + 3}:K{i + 3}"
            write_sheet.range(data_range).value = wriet_data_list[i]
    write_workbook.save(output_file)
    read_workbook.close()
    write_workbook.close()
    app.quit()
    print(f"文件已输出为:{output_file}")
    input("输入任意值退出")
if __name__ == "__main__":
    input_file = "input/pu_input.xls"
    output_file = "output/pu_output.xls"
    pu_to_pu(input_file, output_file)