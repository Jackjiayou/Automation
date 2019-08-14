import  win32com.client  as win32_client
import logging
import PublicData

def call_vba_macro(file_path,func_name,run_type,import_path,AssigneeTaxCostAnalysis_path,Unmapped_path,AccountReallocation_path):
    ''' call vba macro ''' 
    try:
        xlApp = win32_client.Dispatch('Excel.Application')
        xlApp.visible = False # 此行设置打开的Excel表格为可见状态；忽略则Excel表格默认不可见
        #xlApp.DispalyAlerts = 0
        xlBook = xlApp.Workbooks.Open(file_path,False)
        result = xlBook.Application.Run(func_name,run_type,import_path,AssigneeTaxCostAnalysis_path,Unmapped_path,AccountReallocation_path)
        xlBook.Close(False)
        xlApp.quit()
        if result == 'success':
            return True
        else:
            return False
    except Exception as ex:
        xlBook.Close(False)
        xlApp.quit()
        logging.info('call_vba_macro :'+str(ex))
        return False

def run_vba_blackline_macro(file_path,func_names):
    try:
        report_foler = PublicData.download_report_folder_path_total
        import_path = report_foler+'\TransactionDetailReport_ImpotReport_'+PublicData.user_name+'.xlsx'
        AssigneeTaxCostAnalysis_path = report_foler+'\AssigneeTaxCostAnalysis_'+PublicData.user_name+'.xlsx'
        Unmapped_path =report_foler+'\TransactionDetailReport_UnmappReport_'+PublicData.user_name+'.xlsx'
        AccountReallocation_path =report_foler+'\TransactionDetailReport_AccountReallocationReport_'+PublicData.user_name+'.xlsx'
        is_success = call_vba_macro(file_path,func_names,1,import_path,AssigneeTaxCostAnalysis_path,Unmapped_path,AccountReallocation_path)
        return is_success
    except Exception as ex:
        logging.info("get_last_month_shorthand :" + str(ex))
        return False