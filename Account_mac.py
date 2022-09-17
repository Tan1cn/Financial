import pandas as pd
import openpyxl

def strip_in_data(data):  # 把列名中和数据中首尾的空格都去掉。
    data = data.rename(columns={column_name: column_name.strip() for column_name in data.columns})
    data = data.applymap(lambda x: x.strip().strip('¥') if isinstance(x, str) else x)
    return data


ccb1 = pd.read_csv('储蓄卡账单.txt',header=3, skipfooter=0, encoding='utf-8')
ccb1 = ccb1.iloc[:, [0, 4, 7, 1, 2, 3, 5]]  # 按顺序提取所需列
ccb1 = strip_in_data(ccb1)  # 去除列名与数值中的空格。

'''
print(ccb1)

def read_data_ccb1(path):  # 获取建设银行储蓄卡数据
    d_ccb1 = pd.read_csv(path, header=16, skipfooter=0, encoding='utf-8')  # 数据获取，微信
    d_ccb1 = d_ccb1.iloc[:, [0, 4, 7, 1, 2, 3, 5]]  # 按顺序提取所需列
    d_ccb1 = strip_in_data(d_ccb1)  # 去除列名与数值中的空格。
    d_ccb1.iloc[:, 0] = d_ccb1.iloc[:, 0].astype('datetime64')  # 数据类型更改
    d_ccb1.iloc[:, 6] = d_ccb1.iloc[:, 6].astype('float64')  # 数据类型更改
    d_ccb1 = d_ccb1.drop(d_ccb1[d_ccb1['收/支'] == '/'].index)  # 删除'收/支'为'/'的行
    d_ccb1.rename(columns={'当前状态': '支付状态', '交易类型': '类型', '金额(元)': '金额'}, inplace=True)  # 修改列名称
    d_ccb1.insert(1, '来源', "微信", allow_duplicates=True)  # 添加微信来源标识
    len1 = len(d_ccb1)
    print("成功读取 " + str(len1) + " 条「微信」账单数据\n")
    return d_ccb1

ccb1 = pd.read_csv('储蓄卡账单.txt',header=3, skipfooter=0, encoding='utf-8')
ccb1 = strip_in_data(ccb1)  # 去除列名与数值中的空格。
ccb1 = ccb1.iloc[:,[1,3]]
print(ccb1)
