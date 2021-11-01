# -*- coding: utf-8 -*-
# @Time    : 2021/2/5 17:01
# @Author  : KEVINKANG
# @File    : main.py
# @Software: PyCharm

import pandas as pd
from pathlib import Path

#遍历目标文件夹下的所有xlsx文件
files = Path(r"F:\coding\ExcelCombination\汇总").glob("*.xlsx")

#设置一个空的容器
dfs = []

#读取指定位置的
for f in files:

    table = pd.read_excel(f, sheet_name='月结详细信息')  #读取表格的指定工作簿'月结详细信息'
    part = table.iloc[6:]  # 取表格第7行及以后
    dfs.append(part)  #添加列表

#合并列表为一个数列
df = pd.concat(dfs)
#输出为excel
df.to_excel("result2.xlsx", index = False)