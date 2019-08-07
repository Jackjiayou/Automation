import module1
import datetime
import excel_helper
import Utility
import PublicData
import os

approve_reject_list = []
item ={'Country':'2123','Description':'456','Reason':'789'}
item1 ={'Country':'2123','Description':'456','Reason':'789'}
item2 ={'Country':'2123','Description':'456','Reason':'789'}
approve_reject_list.append(item)
approve_reject_list.append(item1)
approve_reject_list.append(item2)
try:
    source_file = PublicData.temp_report_excel_file_path
    target_file = PublicData.blackline_monthly_reconciliation_previous_month_file_path

    if not os.path.isfile(PublicData.blackline_monthly_reconciliation_previous_month_file_path):
        excel_helper.create_excel_file(PublicData.blackline_monthly_reconciliation_previous_month_file_path)
        excel_helper.create_sheet_name(PublicData.blackline_monthly_reconciliation_previous_month_file_path,PublicData.log_approve_sheet_name)
    excel_helper.copy_sheet(PublicData.log_approve_sheet_name,source_file,target_file)
    print(123)
    #if not os.path.isfile(PublicData.blackline_monthly_reconciliation_previous_month_file_path):
    #    excel_helper.create_excel_file(PublicData.blackline_monthly_reconciliation_previous_month_file_path)
    
    #excel_helper.create_sheet_name(PublicData.blackline_monthly_reconciliation_previous_month_file_path,PublicData.log_approve_sheet_name)
    #colunm_names = ['Country','Description','Reason']
    #excel_helper.set_colunm_name(colunm_names,PublicData.blackline_monthly_reconciliation_previous_month_file_path,PublicData.log_approve_sheet_name)

    #for info in approve_reject_list:
    #    arry_info=[info['Country'],info['Description'],info['Reason']]
    #    excel_helper.write_row_data(arry_info,PublicData.blackline_monthly_reconciliation_previous_month_file_path,PublicData.log_approve_sheet_name)
    #for info in approve_reject_list:
    #    arry_info=[info['Country'],info['Description'],info['Reason']]
    #    excel_helper.write_row_data(arry_info,PublicData.blackline_monthly_reconciliation_previous_month_file_path,PublicData.log_approve_sheet_name)
except Exception as ex:
    print(ex)
