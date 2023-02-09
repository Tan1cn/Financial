import pandas as pd
import openpyxl
import datetime
import os
import tkinter.filedialog

#---------------------------------------------------------------------------------------------------

def strip_in_data(data):  # 把列名中和数据中首尾的空格都去掉。
    data = data.rename(columns={column_name: column_name.strip() for column_name in data.columns})
    data = data.applymap(lambda x: x.strip().strip('¥') if isinstance(x, str) else x)
    return data

def read_data_ccb1(path):  # 获取建设银行储蓄卡数据,并调整数据格式
    ccb1 = pd.read_csv(path,header=3, skipfooter=0, encoding='utf-8')
    ccb1 = ccb1.iloc[:, [1, 2, 3, 4, 7, 9, 10]]  # 按顺序提取所需列
    ccb1 = strip_in_data(ccb1)  # 去除列名与数值中的空格。
    ccb1.iloc[:, 0] = ccb1.iloc[:, 0].astype('str')  # 数据类型更改
    ccb1.iloc[:, 0] = ccb1.iloc[:, 0].astype('datetime64')  # 数据类型更改
    ccb1.iloc[:, 1] = ccb1.iloc[:, 1].astype('datetime64')  # 数据类型更改
    ccb1.iloc[:, 2] = ccb1.iloc[:, 2].astype('float64')  # 数据类型更改
    ccb1.iloc[:, 3] = ccb1.iloc[:, 3].astype('float64')  # 数据类型更改
    ccb1.rename(columns={'交易地点': '交易详情'}, inplace=True)  # 修改列名称
    ccb1.insert(7, '数据来源', "建设银行储蓄卡", allow_duplicates=True)  # 添加建设银行储蓄卡来源标识
    ccb1 = ccb1.fillna(0)
    len1 = len(ccb1)
    print("成功读取 " + str(len1) + " 条「建设银行储蓄卡」账单数据\n")
    #print(ccb1)
    #print(ccb1.dtypes)
    return(ccb1)

def read_data_ccb2(path):  # 获取建设银行信用卡数据,并调整数据格式
    ccb2 = pd.read_csv(path,header=3, skipfooter=0, encoding='GB2312',encoding_errors='ignore')
    ccb2 = ccb2.iloc[:, [0, 5, 3, 6]]  # 按顺序提取所需列
    ccb2 = strip_in_data(ccb2)  # 去除列名与数值中的空格。
    ccb2.rename(columns={'交易日': '交易日期', '入账金额': '支出', '类型': '摘要','交易描述': '交易详情'}, inplace=True)  # 修改列名称
    ccb2.insert(1, '交易时间', "00:00:00")  
    ccb2.insert(3, '收入', "")  
    ccb2.insert(5, '对方户名', "无")
    ccb2.insert(7, '数据来源', "建设银行信用卡")# 添加建设银行信用卡来源标识
    ccb2.iloc[:, 0] = ccb2.iloc[:, 0].astype('str')  # 数据类型更改
    ccb2.iloc[:, 0] = ccb2.iloc[:, 0].astype('datetime64')  # 数据类型更改
    ccb2.iloc[:, 1] = ccb2.iloc[:, 1].astype('datetime64')  # 数据类型更改
    ccb2.iloc[:, 2] = ccb2.iloc[:, 2].astype('float64')  # 数据类型更改
    for i in ccb2.index:       #将“支出/收入”单列转换为支出、收入分别一列
        if ccb2.iloc[i,2] < 0:
            ccb2.iloc[i,3]= -ccb2.iloc[i,2]
            ccb2.iloc[i,2]= 0
        else:
            ccb2.iloc[i,3]= 0 
    ccb2.iloc[:, 3] = ccb2.iloc[:, 3].astype('float64')  # 数据类型更改
    #print(ccb2)
    #print(ccb2.dtypes)
    len2 = len(ccb2)
    print("成功读取 " + str(len2) + " 条「建设银行信用卡」账单数据\n") 
    return(ccb2)

def read_data_wx1(path):  # 获取微信数据,并调整数据格式
    wx1 = pd.read_csv(path,header=16, skipfooter=0, encoding='utf_8',encoding_errors='ignore')
    wx1.iloc[:, 5] = wx1.iloc[:, 5].str[1:].astype('float64')  # 数据类型更改
    for i in wx1.index:       #将收入赋负号
        if wx1.iloc[i,4] == '收入':
            wx1.iloc[i,5]= - wx1.iloc[i,5]
    wx1 = wx1[ ~ wx1['支付方式'].str.contains('银行') ]
    wx1 = wx1.iloc[:, [0, 0, 5, 1, 2, 3]]  # 按顺序提取所需列
    wx1 = strip_in_data(wx1)  # 去除列名与数值中的空格。
    wx1.columns = ['交易日期','交易时间','支出','摘要','对方户名','交易详情']
    wx1.insert(3, '收入', "")  
    wx1.insert(7, '数据来源', "微信")# 添加微信来源标识
    wx1.iloc[:, 0] = wx1.iloc[:, 0].astype('str')  # 数据类型更改
    wx1.iloc[:, 0] = wx1.iloc[:, 0].astype('datetime64')  # 数据类型更改
    wx1.iloc[:, 1] = wx1.iloc[:, 1].astype('datetime64')  # 数据类型更改
    wx1.iloc[:, 2] = wx1.iloc[:, 2].astype('float64')  # 数据类型更改
    wx1.reset_index(drop=True, inplace=True)

    for i in wx1.index:       #将“支出/收入”单列转换为支出、收入分别一列
        if wx1.iloc[i,2] < 0:
            wx1.iloc[i,3]= -wx1.iloc[i,2]
            wx1.iloc[i,2]= 0
        else:
            wx1.iloc[i,3]= 0 
    wx1.iloc[:, 3] = wx1.iloc[:, 3].astype('float64')  # 数据类型更改
    #print(wx1)
    #print(wx1.dtypes)
    len2 = len(wx1)
    print("成功读取 " + str(len2) + " 条「微信」账单数据\n") 
    return(wx1)

# def read_data_zfb1(path):

#---------------------------------------------------------------------------------------------------

if __name__ == '__main__':

    print('提示：请在弹窗中选择要导入的【储蓄卡账单】账单文件\n')
    path_cxk1 = tkinter.filedialog.askopenfilename(title='选择要导入的储蓄卡账单：', filetypes=[('所有文件', '.*'), ('csv文件', '.csv')])
    print('提示：请在弹窗中选择要导入的【信用卡账单】账单文件\n')
    path_xyk1 = tkinter.filedialog.askopenfilename(title='选择要导入的信用卡账单：', filetypes=[('所有文件', '.*'), ('csv文件', '.csv')])
    print(path_xyk1)
    print('提示：请在弹窗中选择要导入的【微信账单】账单文件\n')
    path_wx1 = tkinter.filedialog.askopenfilename(title='选择要导入的微信账单：', filetypes=[('所有文件', '.*'), ('csv文件', '.csv')])
    # print('提示：请在弹窗中选择要导入的【支付宝账单】账单文件\n')
    # path_zfb1 = tkinter.filedialog.askopenfilename(title='选择要导入的支付宝账单：', filetypes=[('所有文件', '.*'), ('csv文件', '.csv')])

    print('提示：请在弹窗中选择要保存的账单\n')
    path_account = tkinter.filedialog.askopenfilename(title='选择要保存的账单：', filetypes=[('所有文件', '.*'), ('xlsx文件', '.xlsx')])


    data_ccb1 = data_ccb2 = data_wx1 = path_zfb1 = pd.DataFrame(columns=['交易日期','交易时间','支出','收入','摘要','对方户名','交易详情','数据来源'])
    
    if os.path.exists(path_cxk1):
        data_ccb1 = read_data_ccb1(path_cxk1)
    if os.path.exists(path_xyk1):
        data_ccb2 = read_data_ccb2(path_xyk1)
    if os.path.exists(path_wx1):
        data_wx1 = read_data_wx1(path_wx1)
    # elif os.path.exists(path_zfb1):
    #     data_zfb1 = read_data_zfb1(path_zfb1)        
    print(data_ccb2)
    
    data_merge = pd.concat([data_ccb1,data_ccb2,data_wx1,path_zfb1])
    merge_list = data_merge.values.tolist()
    print("导入以下数据：")
    print(data_merge)
    print("-"*150)

    #---------------------------------------------------------------------------------------------------

    workbook = openpyxl.load_workbook(path_account)  # openpyxl读取账本文件
    sheet = workbook['财务记录']
    maxrow = sheet.max_row  # 获取最大行
    print('\n「明细」 sheet 页已有 ' + str(maxrow) + ' 行数据，将在末尾写入数据')
    for row in merge_list:
        sheet.append(row)  # openpyxl写文件

    workbook.save(path_account)  # 保存
    print("\n成功将数据写入到 " + path_account)
