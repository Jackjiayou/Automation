import openpyxl
import log_helper
from openpyxl import load_workbook

def copy_sheet(sheetname,source_file_path,target_file_path):
 '''copy sheet '''
 source_wb = load_workbook(source_file_path)
 target_wb = load_workbook(target_file_path)
 
 source_ws = source_wb[sheetname]
 target_ws = target_wb[sheetname]
 
 #两个for循环遍历整个excel的单元格内容
 for i,row in enumerate(source_ws.iter_rows()):
  for j,cell in enumerate(row):
   target_ws.cell(row=i+1, column=j+1, value=cell.value)
 
 target_wb.save(target_file_path)

def write_list_data_to_excel(list_data,file_path,sheet_name):
    ''' write list data to excel '''
    wb = openpyxl.load_workbook(file_path)
    try:
        sheet_report = wb.get_sheet_by_name(sheet_name) 
        if len(list_data) > 0:
            for row in list_data:
                sheet_report.append(row)
        return True
    except Exception as ex:
        return False
        log_helper.add_log('set_colunm_name :' + str(ex))
    finally:
        wb.save(file_path)
        wb.close()

def create_excel_file(file_path):
    '''create excel file '''
    try:
        wb = openpyxl.Workbook()
        wb.save(file_path)
        return True
    except Exception as ex:
        log_helper.add_log('create_excel_file :' + str(ex))
        return False

def create_sheet_name(file_path,sheet_name,index = None):
    wb = openpyxl.load_workbook(file_path)
    try:
        sheet_names = wb.get_sheet_names()

        have_sheet = False
        for name in sheet_names:
            if sheet_name == name:
                have_sheet = True
                break
        if not have_sheet:
             wb.create_sheet(title=sheet_name,index = index)
        wb.save(file_path)
        wb.close()
        return True
    except Exception as ex:
        log_helper.add_log('create_sheet_name :' + str(ex))
        return False
        

        
def set_colunm_name(table_title,file_path,sheet_name):
    wb = openpyxl.load_workbook(file_path)
    try:
        #tableTitle = ['userName', 'Phone', 'age', 'Remark']
        ws = wb.get_sheet_by_name(sheet_name) 
        for col in range(len(table_title)):
            c = col + 1
            ws.cell(row=1, column=c).value = table_title[col]
        return True
    except Exception as ex:
        return False
        log_helper.add_log('set_colunm_name :' + str(ex))
    finally:
        wb.save(file_path)
        wb.close()

def write_row_data(data_list,file_path,sheet_name):
    wb = openpyxl.load_workbook(file_path)
    try:
        sheet_report = wb.get_sheet_by_name(sheet_name) 
       
        sheet_report.append(data_list)
        return True
    except Exception as ex:
        return False
        log_helper.add_log('set_colunm_name :' + str(ex))
    finally:
        wb.save(file_path)
        wb.close()

