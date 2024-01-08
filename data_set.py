# coding=utf-8
import subprocess
# 删除excel进程，防止文件被占用。
subprocess.call('taskkill /f /im excel.exe', shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
# EXCEL
excel_range=['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z',
             'AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ']
# 一、dat部分
# 模板文件
dat_to_xls_template_file= "data/dat_to_xls_template.xls"
#报错
crash_list=[]
#控制卡
dat_control_is_chinese_list=[18,21,26]
dat_control_keys_list=[
    'CASEID',
    'PROJECT',
    'NEW_BASE',
    'OLD_BASE',
    'PF_MAP',
    'MVA_BASE',
    'AI_CONTROL',
    'LTC',
    'DECOUPLED',
    'CURRENT',
    'NEWTON',
    'BUSV',
    'AIPOWER',
    'TX',
    'Q',
    'OPCUT',
    'P_INPUT_LIST_MODEL',
    'INPUT_ZONES',
    'ERRORS',
    'P_OUTPUT_LIST_MODEL',
    'OUTPUT_ZONES',
    'FAILED_SOL',
    'RPT_SORT',
    'P_ANALYSIS_RPT_LEVEL',
    'P_ANALYSIS_RPT_AREA',
    'P_ANALYSIS_ZONES',
    'P_ANALYSIS_OWNERS',
    'AI_LIST',
    'OUTPUTDEC',
    'RX_CHECK']
dat_default_value_dict={
    'MVA_BASE':100,
    'DECOUPLED':2,
    'CURRENT':0,
    'NEWTON':30,
    'BUSV':0.005,
    'AIPOWER':0.005,
    'TX':0.005,
    'Q':0.005,
    'OPCUT':0.0001,
    'P_INPUT_LIST':'NONE',
    'ERRORS':'NO_LIST',
    'P_OUTPUT_LIST':'NONE',
    'FAILED_SOL':'FULL_LIST',
    'P_ANALYSIS_RPT1':1,
    'P_ANALYSIS_RPT2':'*',
    'AI_LIST':'FULL',
    'OUTPUTDEC':2,
    'RX_CHECK':'ON'}#暂时没用
#节点数据卡
dat_bus_list=[]
dat_bus_len_format_list=[2,1,3,8,4,2,5,5,4,4,4,5,5,5,4,4]
dat_bus_no_digit_list=[0,1,2,3,5]
#支路数据卡
dat_line_list=[]
dat_line_len_format_list=[2,1,3,8,4,1,8,4,1,1,4,1,6,6,6,6,3]
dat_line_no_digit_list=[0,1,2,3,5,6,8,9]
#变压器数据卡
dat_tran_len_format_list=[2,1,3,8,4,1,8,4,1,1,4,1,6,6,6,6,5,5]
dat_tran_no_digit_list=[0,1,2,3,5,6,8,9]
#B-PQ节点 BQ-无功有限制的PV节点 BE-PV节点 BS-平衡节点
dat_sheet_b_list = ['节点类型',
                '修改码',
                '所有者',
                '节点名称',
                '节点基准电压(KV)',
                '分区名',
                '恒定有功负荷(MW)',
                '恒定无功负荷(MVAR),+表示感性',
                '并联导纳有功负荷(MW)',
                '并联导纳无功负荷(MVAR),+表容性',
                '最大有功出力Pmax(MW)',
                '实际有功出力PGen(MW)',
                '无功出力最大值(MVAR),+表容性',
                '无功出力最小值(MVAR),+表容性',
                '安排的电压值或Vmax(标么值)',
                '安排的Vmin(标么值)']
dat_sheet_l_list = ['线路类型',
                '修改码',
                '所有者',
                '节点1名称',
                '节点1基准电压(KV)',
                '区域联络线测点标志',
                '节点2名称',
                '节点2基准电压(KV)',
                '并联线路的回路标志',
                '分段串连而成的线路的段号',
                '线路额定电流(安培)',
                '并联线路数目',
                '电阻标幺值',
                '电抗标幺值',
                '线路对地电导标么值(G/2)',
                '线路对地电纳标幺值(B/2)',
                '线路或段的长度']
dat_sheet_t_list = ['变压器类型',
                '修改码',
                '所有者',
                '节点1名称',
                '节点1基准电压',
                '区域联络线测点标志',
                '节点2名称',
                '节点2基准电压',
                '并联线路的回路标志',
                '分段串连而成的线路的段号',
                '变压器额定容量(MVA)',
                '并联变压器数目',
                '由铜损引起的等效电阻(标么值)',
                '漏抗(标么值)',
                '由铁损引起的等效电导(标么值)',
                '激磁导纳(标么值)',
                '节点1的固定分接头',
                '节点2的固定分接头']
# 二、swi部分
# 模板文件
swi_to_xls_template_file= "data/swi_to_xls_template.xls"
# 第一个部分填写所有模型、故障操作参数。
# 计算控制CASE卡数据
swi_CASE_is_blank_list = [1,3,5,7,9]
swi_CASE_len_format_list = [4,1,10,1,1,6,1,5,1,14,5,5,5,5,5,5,6]
# 计算控制继续F1卡数据
swi_F1_is_blank_list = [1,3,5,7,9,11,13,15]
swi_F1_len_format_list = [2,2,5,10,3,3,1,1,4,1,1,3,5,9,1,29]
# 监视曲线控制F0卡数据
swi_F0_is_blank_list = [1,3,5,9,13,16,18,21,23]
swi_F0_len_format_list = [2,2,1,2,1,1,8,4,1,1,8,4,1,1,5,5,2,1,1,8,4,2,1,1,8,5]
# 计算控制FF卡数据
swi_FF_is_blank_list = [1,3,5,7,9,11,13,15,17,19,21,23,25,29,30]
swi_FF_len_format_list = [2,6,3,1,5,1,3,1,3,1,5,1,3,1,4,1,2,1,3,7,2,
                          1,1,1,1,3,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1]
# 基本故障模型LS卡数据
swi_LS_is_blank_list = []
swi_LS_len_format_list = []
#  风电机组模型MY卡数据
swi_MY_is_blank_list = [1,5,9]
swi_MY_len_format_list = [2,1,8,4,1,6,3,3,4,25,5,5,6,4,3]
#  风电机组模型下的MR卡——低电压穿越保护模型
swi_MR_is_blank_list = [1,6,11]
swi_MR_len_format_list = [2,1,8,4,1,1,38,5,1,4,4,11]
#  风电机组模型下的EU卡——网侧变频器有功控制模型
swi_EU_is_blank_list = [1,16]
swi_EU_len_format_list = [2,1,8,4,1,5,5,5,5,5,5,5,5,5,5,1,13]
#  风电机组模型下的EZ卡——正常运行状态下无功控制模型
swi_EZ_is_blank_list = [1,14]
swi_EZ_len_format_list = [2,1,8,4,1,5,5,5,5,5,5,5,5,5,1,1,5,8,4]
#  风电机组模型下的ES卡——有功无功电流限制模型
swi_ES_is_blank_list = [1,5,13]
swi_ES_len_format_list = [2,1,8,4,1,9,5,5,5,1,1,1,5,12,5,5,5,5]
#  风电机组模型下的EV卡——低电压高电压状态判断模型
swi_EV_is_blank_list = [1,5,7,14]
swi_EV_len_format_list = [2,1,8,4,1,1,1,1,5,5,5,5,5,5,31]
#  风电机组模型下的LP卡——低电压穿越状态下有功控制模型
swi_LP_is_blank_list = [5,17]
swi_LP_len_format_list = [3,8,4,1,2,1,1,5,4,1,5,5,1,5,5,5,5,19]
#  风电机组模型下的LQ卡——低电压穿越状态下无功控制模型
swi_LQ_is_blank_list = [4,12,16,18]
swi_LQ_len_format_list = [3,8,4,1,2,1,1,5,5,1,5,5,10,1,5,5,8,5,5]
# 读取发电机次暂态参数模型M卡数据
swi_M_is_blank_list = [1,6,8,10,12,17]
swi_M_len_format_list = [1,2,8,4,1,5,1,3,5,2,1,3,1,5,5,4,4,15,5,5]
# 发电机双轴模型数据MF卡数据
# 发电机E'恒定模型数据MC卡数据
swi_MC_or_MF_is_blank_list = [1,]
swi_MC_or_MF_len_format_list = [2,1,8,4,1,6,3,3,4,4,5,5,5,5,4,3,5,5,4,3]
# 静态负荷模型LB卡数据
swi_LB_or_LC_is_blank_list = [1,16]
swi_LB_or_LC_len_format_list = [2,1,8,4,2,10,5,5,5,5,5,5,5,5,5,5,3]
# 第二个部分填写输出数据。
# 输出主控制MH卡数据
swi_MH_is_blank_list = [1,3,6]
swi_MH_len_format_list = [2,1,2,71,1,1,2]
# 母线输出控制BH卡数据
swi_BH_is_blank_list = [1,4]
swi_BH_len_format_list = [2,1,1,1,75]
# 母线输出B卡数据
swi_B_is_blank_list = [1,4,7,9,11,13,15,26]
swi_B_len_format_list = [1,2,8,4,2,1,1,1,1,2,1,2,1,4,3,3,1,1,1,1,1,1,1,1,1,1,33]
# 发电机输出控制GH卡
swi_GH_is_blank_list = [1,3,6,8]
swi_GH_len_format_list = [2,2,1,1,8,4,1,1,60]
# 发电机输出G卡数据
swi_G_is_blank_list = [1,4,6,8,10,12,14,16,18,20,22,24,26,28,30,32,34,39]
swi_G_len_format_list = [1,2,8,4,1,1,2,1,2,1,2,1,2,1,2,1,2,1,2,1,2,1,2,1,2,1,2,1,2,1,2,1,2,1,3,8,4,1,3,1,1]
#三、潮流计算输出pfo文件
# 模板文件
pfo_to_xls_template_file= "data/pfo_to_xls_template.xls"
# 节点数据
# 线路数据
pfo_L_len_format_list = [15,15,6,12,12,12,12,26]
#四、标幺值转换
pu_to_pu_template_file = "data/pu_to_pu_template.xls"
