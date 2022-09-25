'''
@project:TEST_FRAME
@auther:liuyun
@file:read_excel01.py
@date:2022/9/25 16:45
@desc:
'''
import os
import xlrd
from common.excel_utils import ExcelUtils
current_path=os.path.dirname(__file__)#common目录
excel_path=os.path.join(current_path,'../','samples/data/test_excel.xls')
print(excel_path)
excelUtils=ExcelUtils(excel_path,'Sheet1')
print(excelUtils.get_merged_cellvalue(4,0)) # 第5行第1列

# 数据最终从excel要转成一下格式
# {"事件":"学习python编程","步骤序号":step_01,"操作步骤":"购买微课"},
# {"事件": "学习python编程", "步骤序号":step_01, "操作步骤":"购买微课"}
# sheet_list=[]
# for row in range(1,excelUtils.get_row_count()):#0行是标题
#     row_dict={}
#     row_dict['事件']=excelUtils.get_merged_cellvalue(row,0)
#     row_dict['步骤序号'] = excelUtils.get_merged_cellvalue(row, 1)
#     row_dict['步骤操作'] = excelUtils.get_merged_cellvalue(row, 2)
#     row_dict['完成情况'] = excelUtils.get_merged_cellvalue(row, 3)
#     sheet_list.append(row_dict)
# for row in sheet_list:
#     print(row)

sheet_list=[]
for row in range(1,excelUtils.get_row_count()):#0行是标题  代码盖章
    row_dict={}
    row_dict['事件']=excelUtils.get_merged_cellvalue(row,0)
    row_dict['步骤序号'] = excelUtils.get_merged_cellvalue(row, 1)
    row_dict['步骤操作'] = excelUtils.get_merged_cellvalue(row, 2)
    row_dict['完成情况'] = excelUtils.get_merged_cellvalue(row, 3)
    sheet_list.append(row_dict)

first_row=excelUtils.sheet.row(0) #第一行
for row in range(1,excelUtils.get_row_count()):
    row_dict = {}
    for col in range(1,excelUtils.get_col_count()):
        row_dict[first_row[col].value]=excelUtils.get_merged_cellvalue(row,col)
        sheet_list.append(row_dict)

for row in sheet_list:
    print(row)
