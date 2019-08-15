#!/usr/bin/env python    # -*- coding: utf-8 -*
import os
import sys
import getpass
import datetime

username = ''
password = ''
#get time 
today = datetime.date.today()
first = today.replace(day=1)
last_month = first - datetime.timedelta(days=1)
previous_month_number = last_month.month
download_report_region_name = 'EALA'
#app folder pervoius month name
month_list =['January','February','March','April','May','June','July','August','September','October','November','December']
app_previous_month_folder_name =month_list[previous_month_number-1]

#report in BlacklineMonthlyReconciliation folder file path
month_list_for_blackline_monthly_reconciliation_report =['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
report_name_in_monthly_reconciliation_folder ='Report_'+str( month_list_for_blackline_monthly_reconciliation_report[previous_month_number-1])+str(last_month.year)[-2:]+'.xlsx'
blackline_monthly_reconciliation_previous_month_file_path =r'C:\Blackline Automation App\Report'+'\\'+app_previous_month_folder_name+'\\'+download_report_region_name+r'\BlacklineMonthlyReconciliation\\'+report_name_in_monthly_reconciliation_folder
eala_foder_log_excel_path = r'C:\Blackline Automation App\Report'+'\\'+app_previous_month_folder_name+'\\'+'EALA'+r'\BlacklineMonthlyReconciliation\\'+report_name_in_monthly_reconciliation_folder
#config download report sheet name
config_sheet_name ='Download Config'
upload_report_folder = r'C:\Blackline Automation App\Report'+'\\'+app_previous_month_folder_name+'\\'+download_report_region_name+r'\ImportGroup'

upload_eale_file_foler = r'C:\Blackline Automation App\Report'+'\\'+app_previous_month_folder_name+'\\EALA\\ImportGroup'
upload_apac_na_file_foler = r'C:\Blackline Automation App\Report'+'\\'+app_previous_month_folder_name+'\\APAC&NA\\ImportGroup'


upload_and_approve_excel_log_file_path = r'C:\Blackline Automation App\Report'+'\\'+app_previous_month_folder_name+'\\'+download_report_region_name+r'\BlacklineMonthlyReconciliation\Report_May19.xlsx'
app_log_file_path = os.path.expanduser('~') +r'\Blackline Automation App\Log\err.log'
logFileDir =  os.path.expanduser('~')+r'\Blackline Automation App\Log'
currentUserDir = os.path.expanduser('~')+r"\AppData\Local\Google\Chrome\User Data"
myRootPath = sys.path[0].replace('base_library.zip','')
task_config_file_path = myRootPath+r'\Resource\DownloadGtfReportConfig.xlsx'
web_driver = None
browser_type ='chrome'
download_report_folder_path = r'C:\Blackline Automation App\Report'+'\\'+app_previous_month_folder_name+'\\'+download_report_region_name+r'\Source File_'+ username
download_report_folder_path_total =  r'C:\Blackline Automation App\Report'+'\\'+app_previous_month_folder_name+'\\'+download_report_region_name+r'\Source File'
#os.path.expanduser('~') +r'\Blackline Automation App\Source File'
generate_report_blackline_folder_path_by_vba_tool =r'C:\Blackline Automation App\Report'+'\\'+app_previous_month_folder_name+'\\'+download_report_region_name+r'\BlacklineMonthlyReconciliation'
generate_report_import_folder_path_by_vba_tool =r'C:\Blackline Automation App\Report'+'\\'+app_previous_month_folder_name+'\\'+download_report_region_name+r'\ImportGroup'
my_root_tool_path = myRootPath+r'\Resource\EALA\tool.xlsm'
vba_tool_folder = r'C:\Blackline Automation App\Report'+'\\'+app_previous_month_folder_name+'\\'+download_report_region_name
my_root_template_folder = myRootPath+r'\Resource\EALA\Template'
template_folder = r'C:\Blackline Automation App\Report'+'\\'+app_previous_month_folder_name+'\\'+download_report_region_name+r'\Template'
my_root_template_MonthlyReconReportTemplate_file_path =myRootPath+r'\Resource\EALA\Template\Monthly Recon Report_Template.xlsx'
my_root_template_ReportTemplate_file_path =myRootPath+r'\Resource\EALA\Template\Report_Template.xlsx'

template_MonthlyReconReportTemplate_file_path = r'C:\Blackline Automation App\Report'+'\\'+app_previous_month_folder_name+'\\'+download_report_region_name+r'\Template\Monthly Recon Report_Template.xlsx'
template_ReportTemplate_file_path = r'C:\Blackline Automation App\Report'+'\\'+app_previous_month_folder_name+'\\'+download_report_region_name+r'\Template\Report_Template.xlsx'

temp_report_excel_file_path = 'C:\\Blackline Automation App\Report'+'\\'+app_previous_month_folder_name+r'\Temp Report.xlsx'
download_report_excel_sheet_name = 'Download Report'
# write all log  to temp excel log
upload_and_approve_excel_log_file_path = temp_report_excel_file_path
new_current_download_report_file_path =''
old_current_download_report_file_path =''
download_type=''
str_excel_config_colunm_report_name ='Report Name'
str_excel_config_colunm_default_report_name ='Default Report Name'
str_excel_config_colunm_new_report_name = 'New Report Name'
str_excel_config_colunm_redownload_count = 'Max download Count'
upload_report_need_prepare_count = 0
approve_report_need_prepare_count = 0
prepare_failed_list = []
approve_time = 5
upload_time = 5
log_upload_sheet_name ='Report For Import'
log_approve_sheet_name ='Report For Approver'
gu_list = []
country_list = []
vba_tool_import_macro_name = 'tool.xlsm!importReport'
vba_tool_generate_macro_name = 'tool.xlsm!Module1.generate'
upload_task_config_sheet_name ='Upload Task'
approve_task_config_sheet_name ='Approve Task'
MainFrom = None

#current user name
user_name = getpass.getuser()

#get download report region
region_config_sheet_name = 'Region Config'
region_config_list = []
region_config_column_name_region = 'Region'
region_config_column_name_gu = 'GU'
region_config_column_name_country_name = 'Country Name'
report_sheetname = 'Report'
def set_user_name():
    user_name = username
