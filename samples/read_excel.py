'''
@project:TEST_FRAME
@auther:liuyun
@file:read_excel.py
@date:2022/9/25 11:58
@desc:
'''
import os
import xlrd
excel_path=os.path.join(os.path.dirname(__file__),'data/test_excel.xls')
print(excel_path)
# excel_path='D:/PythonCode/TEST_FRAME/samples/data/test_excel.xls'
wb=xlrd.open_workbook(excel_path,formatting_info=True)#工作薄 formatting_info加上才能读取合并单元格
sheet=wb.sheet_by_name("Sheet1")#表格对象
# sheet=wb.sheet_by_index(0)#表格对象
cell_value=sheet.cell_value(2,2)#直接取值 第三行第三列 行列从0开始的
print(cell_value)
# print("合并单元格左上角第一个值：%s"%str(sheet1.cell_value(1,0)),"合并单元格的其他值会是空：%s"%sheet1.cell_value(2,0))
#处理合并单元格 非第一个单个格问题
# print(sheet1.merged_cells) #返回一个列表 [1,5,0,1] [起始行，结束行，起始列，结束列] 包头不包尾
#逻辑 凡是在merged_cells属性范围内的单元格，值等于左上角单元格
row_index=1;col_index=0
merged=sheet.merged_cells
print(merged)
for (rlow,rhigh,clow,chigh) in merged:#[(1, 5, 0, 1)]
    if(row_index>=rlow and row_index <rhigh):
        if(col_index>=clow and col_index<chigh):
            cell_value=sheet.cell_value(rlow,clow)
            print('row_index:{} col_index:{}  cell_value:{}'.format(row_index,col_index,cell_value))

def get_merged_cellvalue(row_index,col_index):#取合并的单元格的值
    cell_value=None
    for (rlow, rhigh, clow, chigh) in merged:  # [(1, 5, 0, 1)]
        if (row_index >= rlow and row_index < rhigh):
            if (col_index >= clow and col_index < chigh):
                cell_value = sheet.cell_value(rlow, clow)
                # print('row_index:{} col_index:{}  cell_value:{}'.format(row_index, col_index, cell_value))
    return cell_value
# print(get_merged_cellvalue(0,1))#这个是非合并的单元格，不能取，要优化

#来做优化Excel表格 不管是合并还是非合并的单元格都可以取到值
print(get_merged_cellvalue(3,0))
def get_merged_cellvalue(row_index,col_index):#取合并的单元格的值
    cell_value=None
    for (rlow, rhigh, clow, chigh) in merged:  # [(1, 5, 0, 1)]
        if (row_index >= rlow and row_index < rhigh):
            if (col_index >= clow and col_index < chigh):
                cell_value = sheet.cell_value(rlow, clow)
                break; #防止循环进行判断，出现值覆盖的情况
                # print('row_index:{} col_index:{}  cell_value:{}'.format(row_index, col_index, cell_value))
            else:
                cell_value=sheet.cell_value(row_index,col_index)
        else:
            cell_value = sheet.cell_value(row_index, col_index)
    return cell_value

print(get_merged_cellvalue(0,0))
print(get_merged_cellvalue(3,1))


