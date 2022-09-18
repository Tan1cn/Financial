
import pandas as pd
import numpy as np
from datetime import datetime



path = '信用卡账单.csv'
ccb2 = pd.read_csv(path,header=3, skipfooter=0, encoding='GB2312')
ccb2 = ccb2.iloc[:, [0, 5, 5, 3, 6]]  # 按顺序提取所需列
ccb2 = strip_in_data(ccb2)  # 去除列名与数值中的空格。
print(ccb2.dtypes)
ccb2.iloc[:, 0] = ccb2.iloc[:, 0].astype('str')  # 数据类型更改
#for i in ccb2.index:
    #ccb2['交易日期'][i] = datetime.strptime(ccb2['交易日期'][i], '%Y%m%d').strftime('%m/%d/%Y')# 将交易日期改成pandas可识别的日期格式
ccb2.iloc[:, 0] = ccb2.iloc[:, 0].astype('datetime64')  # 数据类型更改
ccb2.iloc[:, 1] = ccb2.iloc[:, 1].astype('datetime64')  # 数据类型更改
ccb2.iloc[:, 2] = ccb2.iloc[:, 2].astype('float64')  # 数据类型更改
ccb2.iloc[:, 3] = ccb2.iloc[:, 3].astype('float64')  # 数据类型更改
len1 = len(ccb2)

ccb2.insert(7, '来源', "建设银行储蓄卡", allow_duplicates=True)  # 添加建设银行储蓄卡来源标识
print("成功读取 " + str(len1) + " 条「建设银行储蓄卡」账单数据\n")
print(ccb2)
print(ccb2.dtypes)