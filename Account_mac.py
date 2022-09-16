import pandas as pd
import openpyxl

def strip_in_data(data):  # 把列名中和数据中首尾的空格都去掉。
    data = data.rename(columns={column_name: column_name.strip() for column_name in data.columns})
    data = data.applymap(lambda x: x.strip().strip('¥') if isinstance(x, str) else x)
    return data

def read_data_ccbcxk(path):  # 获取建设银行储蓄卡数据
    d_wx = pd.read_csv(path, header=16, skipfooter=0, encoding='utf-8')  # 数据获取，微信
    d_wx = d_wx.iloc[:, [0, 4, 7, 1, 2, 3, 5]]  # 按顺序提取所需列
    d_wx = strip_in_data(d_wx)  # 去除列名与数值中的空格。
    d_wx.iloc[:, 0] = d_wx.iloc[:, 0].astype('datetime64')  # 数据类型更改
    d_wx.iloc[:, 6] = d_wx.iloc[:, 6].astype('float64')  # 数据类型更改
    d_wx = d_wx.drop(d_wx[d_wx['收/支'] == '/'].index)  # 删除'收/支'为'/'的行
    d_wx.rename(columns={'当前状态': '支付状态', '交易类型': '类型', '金额(元)': '金额'}, inplace=True)  # 修改列名称
    d_wx.insert(1, '来源', "微信", allow_duplicates=True)  # 添加微信来源标识
    len1 = len(d_wx)
    print("成功读取 " + str(len1) + " 条「微信」账单数据\n")
    return d_wx

ccb1 = pd.read_csv('储蓄卡账单.txt',header=3, skipfooter=0, encoding='utf-8')
ccb1 = strip_in_data(ccb1)  # 去除列名与数值中的空格。
ccb1 = ccb1.iloc[:,[1,3]]
print(ccb1)
