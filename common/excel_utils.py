'''
@project:TEST_FRAME
@auther:liuyun
@file:excel_utils.py
@date:2022/9/25 16:19
@desc:
'''
#导包顺序 内置模块、第三方模块、自定义模块
import os
import xlrd


class ExcelUtils:
    def __init__(self,file_path,sheet_name):
        self.file_path=file_path
        self.sheet_name=sheet_name
        self.sheet=self.get_sheet()#python 支持构造里加普通方法
    def get_sheet(self):
        wb=xlrd.open_workbook(self.file_path,formatting_info=True)#读合并信息是这个参数要加
        sheet=wb.sheet_by_name(self.sheet_name)
        return sheet
    def get_row_count(self):#行数
        row_count=self.sheet.nrows
        return row_count
    def get_col_count(self):#列数
        col_count=self.sheet.ncols
        return col_count
    def get_merged_info(self):
        merged_info=self.sheet.merged_cells
        return merged_info
    def get_cell_value(self,row_index,col_index):
        cell_value=self.sheet.cell(row_index,col_index)
        return cell_value
    def get_sheet_data_by_diact(self):
        all_data_list=[]
        first_row = self.sheet.row(0)  # 获取首行数据
        for row in range(1, excelUtils.get_row_count()):
            row_dict = {}
            for col in range(1, self.get_col_count()):
                row_dict[first_row[col].value] = self.get_merged_cellvalue(row, col)
                all_data_list.append(row_dict)
        return all_data_list
    #取数据的封装
    def get_merged_cellvalue(self,row_index, col_index):  # 取合并的单元格的值
        cell_value = None
        for (rlow, rhigh, clow, chigh) in self.get_merged_info():  # [(1, 5, 0, 1)]
            if (row_index >= rlow and row_index < rhigh):
                if (col_index >= clow and col_index < chigh):
                    cell_value = self.get_cell_value(rlow,clow)
                    break;  # 防止循环进行判断，出现值覆盖的情况
                    # print('row_index:{} col_index:{}  cell_value:{}'.format(row_index, col_index, cell_value))
                else:
                    cell_value = self.get_cell_value(rlow,clow)
            else:
                cell_value = self.get_cell_value(rlow,clow)
        return cell_value

if __name__=='__main__':
    current_path=os.path.dirname(__file__)#common目录
    excel_path=os.path.join(current_path,'../','samples/data/test_excel.xls')
    print(excel_path)
    excelUtils=ExcelUtils(excel_path,'Sheet1')
    print(excelUtils.get_merged_cellvalue(3,0)) # 第5行第1列

    # 数据最终从excel要转成一下格式
    # {"事件":"学习python编程","步骤序号":step_01,"操作步骤":"购买微课"},
    # {"事件": "学习python编程", "步骤序号":step_01, "操作步骤":"购买微课"}

    # 测试获取整行数据
    print(excelUtils.get_sheet_data_by_diact())
