# PSD_BPA卡片读取与写入
## 一、项目描述
这是一个用于文本与表格进行转换的 python 项目，旨在解决电气软件PSD_BPA的稳态数据文件与暂态数据文件的读取与写入问题。
## 二、文件描述
crashreport目录下存放保存文件；  
data目录下存放模板表格文件；  
input目录下存放输入文件；  
output目录下存放输出文件；  
dat_to_xls.py用于将稳态数据文件转换为表格文件；  
swi_to_xls.py用于将暂态数据文件转换为表格文件；  
pfo_to_xls.py用于将潮流计算结果文件转换为表格文件；  
xls_to_dat.py用于将表格文件转换为稳态数据文件；  
dat_datmain.py用于将稳态数据文件与表格文件相互转换；  
pu_to_pu.py用于将额定容量下的标幺值转换为基准容量下的标幺值；  
gen_to_wp.py用于将同步电源转换为直驱发电机机组；  
data-set.py用于设置程序关于卡片格式的初始数据；  
requirements.txt用于列出项目依赖库及其版本的文本文件；  
可用“pip install -r requirements.txt”安装依赖。
## 三、代码描述
程序的输入与输出文件在每个程序末尾的“if __name__ == "__main__":”下可以任意调节，但请注意文件格式，而且表格输入文件请使用data目录下对应的模板表格文件。
## 四、第三方库
xlwings——用于表格文件的读取、写入与修改。
## 五、仓库地址
https://github.com/Arrowss544/PSD_BPA.git
