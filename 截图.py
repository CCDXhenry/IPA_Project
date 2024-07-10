import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.styles import *
import os

# 读取Excel文件中名为'分区县'的工作表
excel1 = pd.read_excel("F:\\IPA_project\\data\\新入网辅助证件统计表6.30.xlsx", sheet_name='分区县', engine='openpyxl')
excel2 = pd.read_excel("F:\\IPA_project\\data\\新入网辅助证件统计表按月.xlsx", sheet_name='分区县', engine='openpyxl')
excel3 = pd.read_excel("F:\\IPA_project\\data\\新入网辅助证件统计表6.30.xlsx", sheet_name='分区县', engine='openpyxl')

# 筛选出列名为'区县'的值为'全通路'的行
filtered_excel1 = excel1[excel1['区县'] == '全通路']
filtered_excel2 = excel2[excel2['区县'] == '全通路']
filtered_excel3 = excel3[excel3['区县'] == '全通路']
excel1_value = filtered_excel1['上传率'].iloc[0]
excel2_value = filtered_excel2['上传率'].iloc[0]
excel3_value = filtered_excel3['上传率'].iloc[0]

# 将'上传率'列的百分比转换为数值
excel2['上传率'] = pd.to_numeric(excel2['上传率'].str.replace('%', ''), errors='coerce') / 100

# 选择'上传率'大于80%的行
filtered_excel2 = excel2[excel2['上传率'] > 0.8]

# 提取符合条件的'区县'列的值
districts = filtered_excel2['区县'].tolist()
print(districts)