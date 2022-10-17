import pandas as pd
import openpyxl
import datetime
import os

path1 = r'C:\Users\tanyi\Downloads\储蓄卡账单.txt'
path2 = r'C:\Users\tanyi\Downloads\信用卡账单.csv'
#path_account = r'D:\Python\Financial\个人财务管理.xlsx'
path_account = r'D:\@生活\台帐\个人财务管理.xlsx'
path_merge = r'本次写入数据.csv'

#---------------------------------------------------------------------------------------------------

def strip_in_data(data):  # 把列名中和数据中首尾的空格都去掉。
    data = data.rename(columns={column_name: column_name.strip() for column_name in data.columns})
    data = data.applymap(lambda x: x.strip().strip('¥') if isinstance(x, str) else x)
    return data


def read_data_ccb1(path):  # 获取建设银行储蓄卡数据
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
    #for i in ccb1.index:       #将收入中的空值转换为0
    #    if bool(pd.isnull(ccb1.iloc[i,3])):
    #        ccb1.iloc[i,3]= 0

    len1 = len(ccb1)
    print("成功读取 " + str(len1) + " 条「建设银行储蓄卡」账单数据\n")
    #print(ccb1)
    #print(ccb1.dtypes)
    return(ccb1)

def read_data_ccb2(path):  # 获取建设银行信用卡数据
    ccb2 = pd.read_csv(path,header=3, skipfooter=0, encoding='GB2312')
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



data_ccb1 = read_data_ccb1(path1)
data_ccb2 = read_data_ccb2(path2)

data_merge = pd.concat([data_ccb1,data_ccb2])
print(data_merge)
merge_list = data_merge.values.tolist()

data_merge.to_csv(path_merge)



workbook = openpyxl.load_workbook(path_account)  # openpyxl读取账本文件
sheet = workbook['财务记录']
maxrow = sheet.max_row  # 获取最大行
print('\n「明细」 sheet 页已有 ' + str(maxrow) + ' 行数据，将在末尾写入数据')
for row in merge_list:
    sheet.append(row)  # openpyxl写文件

workbook.save(path_account)  # 保存
print("\n成功将数据写入到 " + path_account)
