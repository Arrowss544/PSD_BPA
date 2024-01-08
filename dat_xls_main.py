# coding=utf-8
import os
import subprocess #执行命令行用
from xls_to_dat import xls_to_dat
from dat_to_xls import dat_to_xls
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
#获取输入文件
def get_input_file(model):
    excel_sample = "input/039bpa.xls"
    dat_sample = "input/039bpa.DAT"
    if model == 1:
        suffix=".xls"
        print("请选择input目录下的EXCEL文件。")
        print(f"直接回车将采用默认文件:{excel_sample}")
    else:
        suffix = ".DAT"
        print("请选择input目录下的.DAT文件。")
        print(f"直接回车将采用默认文件:{dat_sample}")
    input_files = os.listdir('input/')
    file_choice_list = []
    for i in input_files:
        if i[-4:] == suffix:
            file_choice_list.append(i)
    for i in range(0, len(file_choice_list) + 1):
        if i < len(file_choice_list):
            print(f"{i + 1}.{file_choice_list[i]}")
        else:
            print(f"{i + 1}.退出")
    input_file = input("请输入数字:")
    if is_blank(input_file):
        if model == 1:
            return excel_sample
        else:
            return dat_sample
    elif is_integer_number(input_file):
        n = int(input_file)
        if n == len(file_choice_list)+1:
            return "error"
        else:
            return "input/"+file_choice_list[n-1]
    else:
        print("输入错误！请重新输入！")
        get_input_file(model)
#获取输出文件
def get_output_file(model):
    if model == 1:
        print("请输入.DAT文件的前缀，将保存在output目录下。")
    else:
        print("请输入EXCEL文件的前缀，将保存在output目录下。")
    print("直接回车将与输入文件同名。")
    output_file = input("请输入:")
    if is_blank(output_file):
        return "same"
    else:
        if model == 1:
            return "output/"+output_file+".DAT"
        else:
            return "output/"+output_file+".xls"
#主函数
def main():
    # 删除excel进程，防止文件被占用。
    subprocess.call('taskkill /f /im excel.exe', shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    if os.path.exists("output"):
        pass
    else:
        os.mkdir("output")
    if os.path.exists("input"):
        pass
    else:
        os.mkdir("input")
    print("----------第一步----------")
    print("请选择: \n1.利用EXCEL文件生成.DAT文件。\n2.利用.DAT文件生成EXCEL文件。\n3.退出。")
    choice=input("请输入数字:")
    if int(choice) not in [1,2,3]:
        print("输入错误！")
        main()
    elif int(choice) == 3:
        return
    elif int(choice) == 1:
        print("----------第二步----------")
        input_file = get_input_file(1)
        if input_file == "error":
            main()
        else:
            print("----------第三步----------")
            output_file = get_output_file(1)
            if output_file == "same":
                output_file = "output/"+input_file[6:-3]+"DAT"
            xls_to_dat(input_file, output_file)
            print(f"文件已输出为:{output_file}")
            input("输入任意值退出")
    elif int(choice) == 2:
        print("----------第二步----------")
        input_file = get_input_file(2)
        if input_file == "error":
            main()
        else:
            print("----------第三步----------")
            output_file = get_output_file(2)
            if output_file == "same":
                output_file = "output/" + input_file[6:-3] + "xls"
            try:
                dat_to_xls(input_file, output_file)
            except:
                pass
if __name__ == "__main__":
    main()