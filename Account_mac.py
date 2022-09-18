import pandas as pd
import openpyxl
from datetime import datetime


def strip_in_data(data):  # 把列名中和数据中首尾的空格都去掉。
    data = data.rename(columns={column_name: column_name.strip() for column_name in data.columns})
    data = data.applymap(lambda x: x.strip().strip('¥') if isinstance(x, str) else x)
    return data

def read_data_ccb1(path):  # 获取建设银行储蓄卡数据
    ccb1 = pd.read_csv(path,header=3, skipfooter=0, encoding='utf-8')
    ccb1 = ccb1.iloc[:, [1, 2, 3, 4, 7, 9, 10]]  # 按顺序提取所需列
    ccb1 = strip_in_data(ccb1)  # 去除列名与数值中的空格。
    ccb1.iloc[:, 0] = ccb1.iloc[:, 0].astype('str')  # 数据类型更改
    for i in ccb1.index:
        ccb1['交易日期'][i] = datetime.strptime(ccb1['交易日期'][i], '%Y%m%d').strftime('%m/%d/%Y')# 将交易日期改成pandas可识别的日期格式
    ccb1.iloc[:, 0] = ccb1.iloc[:, 0].astype('datetime64')  # 数据类型更改
    ccb1.iloc[:, 1] = ccb1.iloc[:, 1].astype('datetime64')  # 数据类型更改
    ccb1.iloc[:, 2] = ccb1.iloc[:, 2].astype('float64')  # 数据类型更改
    ccb1.iloc[:, 3] = ccb1.iloc[:, 3].astype('float64')  # 数据类型更改
    len1 = len(ccb1)

    ccb1.insert(7, '来源', "建设银行储蓄卡", allow_duplicates=True)  # 添加建设银行储蓄卡来源标识

    print("成功读取 " + str(len1) + " 条「建设银行储蓄卡」账单数据\n")
    print(ccb1)
    print(ccb1.dtypes)

read_data_ccb1('储蓄卡账单.txt')