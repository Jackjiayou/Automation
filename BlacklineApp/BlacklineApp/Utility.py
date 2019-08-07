# -*- coding: utf-8 -*
#!/usr/bin/env python   
import sys
import os
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import logging
import PublicData
import SeleniumHelper
import  pandas  as pd
import datetime
import time
import pyautogui
from return_value_model import return_data as result
import excel_helper
import shutil
import log_helper
import win32api
import win32con

def remove_file():
    '''remove file'''
    try:
        if os.path.exists(PublicData.new_current_download_report_file_path):
            os.remove(PublicData.new_current_download_report_file_path)
        if os.path.exists(PublicData.old_current_download_report_file_path):
            os.remove(PublicData.old_current_download_report_file_path)
    except Exception as ex:
        logging.exception('remove_file :'+str(ex))
        raise

def copy_log_info(sheet_name):
    ''' copy log from total excel file '''
    try:
        source_file = PublicData.temp_report_excel_file_path
        target_file = PublicData.blackline_monthly_reconciliation_previous_month_file_path
        #if not os.path.isfile(PublicData.blackline_monthly_reconciliation_previous_month_file_path):
        #    excel_helper.create_excel_file(target_file)
        #excel_helper.create_sheet_name(target_file,PublicData.report_sheetname,0)
        #colunm_names = ['Country','Import Total','Unmap Total','Account Reallocation Total','ATCA Total','Variance Checking']
        #excel_helper.set_colunm_name(colunm_names,target_file,PublicData.report_sheetname)
        #excel_helper.create_sheet_name(target_file,PublicData.download_report_excel_sheet_name,index = 1)
        #excel_helper.create_sheet_name(target_file,PublicData.log_approve_sheet_name,index = 2)
        #excel_helper.create_sheet_name(target_file,PublicData.log_upload_sheet_name,index = 3)
        excel_helper.copy_sheet(sheet_name,source_file,target_file)
    except Exception as ex:
        logging('copy_log_info : '+str(ex))

def check_vba_tool_need_files():
    '''check vba tool need file ''' 
    try:
        report_foler = PublicData.download_report_folder_path
        import_path = report_foler+'\TransactionDetailReport_ImpotReport_'+PublicData.user_name+'.xlsx'
        AssigneeTaxCostAnalysis_path = report_foler+'\AssigneeTaxCostAnalysis_'+PublicData.user_name+'.xlsx'
        Unmapped_path =report_foler+'\TransactionDetailReport_UnmappReport_'+PublicData.user_name+'.xlsx'
        AccountReallocation_path =report_foler+'\TransactionDetailReport_AccountReallocationReport_'+PublicData.user_name+'.xlsx'
        
        if not os.path.isfile(import_path) or not os.path.isfile(AssigneeTaxCostAnalysis_path) or not os.path.isfile(Unmapped_path) or not os.path.isfile(AccountReallocation_path)  :
            return False
        else:
            return True
    except Exception as ex:
        logging.exception('check_vba_tool_need_file_path' + str(ex))

def init_app_need_file():
    ''' Initialize app file '''
    try: 
        create_folder(PublicData.generate_report_blackline_folder_path_by_vba_tool)
        create_folder(PublicData.generate_report_import_folder_path_by_vba_tool)
        have_file_log_file = create_file(PublicData.app_log_file_path)
        have_download_report_folder = create_folder(PublicData.download_report_folder_path)
        create_folder(PublicData.template_folder)
        if not os.path.isfile(PublicData.template_MonthlyReconReportTemplate_file_path):
            copy_file(PublicData.my_root_template_MonthlyReconReportTemplate_file_path,PublicData.template_folder)

        if not os.path.isfile(PublicData.template_ReportTemplate_file_path):
            copy_file(PublicData.my_root_template_ReportTemplate_file_path,PublicData.template_folder)

        if not os.path.isfile(PublicData.temp_report_excel_file_path):
            excel_helper.create_excel_file(PublicData.temp_report_excel_file_path)
        excel_helper.create_sheet_name(PublicData.temp_report_excel_file_path,PublicData.download_report_excel_sheet_name,index = 0)
        excel_helper.create_sheet_name(PublicData.temp_report_excel_file_path,PublicData.log_approve_sheet_name,index = 1)
        excel_helper.create_sheet_name(PublicData.temp_report_excel_file_path,PublicData.log_upload_sheet_name,index = 2)

        target_file = PublicData.blackline_monthly_reconciliation_previous_month_file_path
        if not os.path.isfile(target_file):
            excel_helper.create_excel_file(target_file)
        excel_helper.create_sheet_name(target_file,PublicData.report_sheetname,0)
        colunm_names = ['Country','Import Total','Unmap Total','Account Reallocation Total','ATCA Total','Variance Checking']
        excel_helper.set_colunm_name(colunm_names,target_file,PublicData.report_sheetname)
        excel_helper.create_sheet_name(target_file,PublicData.download_report_excel_sheet_name,index = 1)
        excel_helper.create_sheet_name(target_file,PublicData.log_approve_sheet_name,index = 2)
        excel_helper.create_sheet_name(target_file,PublicData.log_upload_sheet_name,index = 3)
        
        if have_file_log_file and have_download_report_folder:
            #logging.info('init_app_need_file_path :  '+ 'have_file_log_file:'+str(have_file_log_file)+'      have_download_report_folder:'+str(have_download_report_folder))
            copy_file(PublicData.my_root_tool_path,PublicData.vba_tool_folder)
            return True
        else:
            return False
        logging.info('end init_app_need_file')
    except Exception as ex:
        logging.exception('init_app_need_file_path error, msg:'+str(ex))
        return False

def generate_download_report(download_report_list):
    try:
        if not os.path.exists(PublicData.temp_report_excel_file_path):
        #    os.remove(PublicData.download_report_excel_file_path)
        #else:
            excel_helper.create_excel_file(PublicData.temp_report_excel_file_path)
        if len(download_report_list)>0:
            excel_helper.create_sheet_name(PublicData.temp_report_excel_file_path,PublicData.download_report_excel_sheet_name,0)
            colunm_names = ['','Report name','Status','Info']
            excel_helper.set_colunm_name(colunm_names,PublicData.temp_report_excel_file_path,PublicData.download_report_excel_sheet_name)
            arry_info=['Log time :'+time.strftime('%Y.%m.%d-%H:%M:%S',time.localtime(time.time())),'','','']
            excel_helper.write_row_data(arry_info,PublicData.temp_report_excel_file_path,PublicData.download_report_excel_sheet_name)    
            for info in download_report_list:
                arry_info=['',info['Report name'],info['Status'],info['Info']]
                excel_helper.write_row_data(arry_info,PublicData.temp_report_excel_file_path,PublicData.download_report_excel_sheet_name)    
    except Exception as ex:
        logging.exception('generate_download_report :' +str(ex))  
        raise

def copy_file(file_path,folder_path):
    try:
        shutil.copy(file_path,folder_path)
    except Exception as ex:
        logging.exception('copy_file :'+ str(ex))
        raise

def create_folder(fileDir):
    '''create folder if not exist '''
    try:
        if not os.path.exists(fileDir):
            os.makedirs(fileDir)
        return True
    except Exception as ex:
        logging.exception('create_folder :'+str(ex))
        return False

def create_file(file_path):
    '''create file if not exist '''
    try:
        fileDir = os.path.dirname(file_path)
        if not os.path.exists(fileDir):
            os.makedirs(fileDir)
        try:
            f =open(file_path,'r')
            f.close()
            return True
        except IOError:
            f = open(file_path,'w')
            f.close()
            return True
    except Excption as ex:
        return False

def create_file_para(filepath,fileDir):
    """ create file if the file not exsited      
        Args:
            filepath: file path
            fileDir: folder path
            
        Returns:

        Raises:
            IOError: An error occurred.
    """
    try:
        if not os.path.exists(fileDir):
            os.makedirs(fileDir)
        try:
            f =open(filepath,'r')
            f.close()
        except IOError:
            f = open(filepath,'w')
            f.close()
    except Exception as ex:
        logging.exception(ex) 
        raise     

def get_report_list(filePath,sheet_name):
    """get report list"""
    try:
        df=pd.read_excel(filePath,sheet_name=sheet_name)
        list_report_config_data=[]
        for i in df.index.values:
            row_data=df.ix[i,[PublicData.str_excel_config_colunm_report_name,PublicData.str_excel_config_colunm_new_report_name,PublicData.str_excel_config_colunm_default_report_name,PublicData.str_excel_config_colunm_redownload_count]].to_dict()
            list_report_config_data.append(row_data)      
        return list_report_config_data
    except Exception as ex:
        logging.exception(ex)

def get_download_report_names(filePath,sheet_name):
    """get gu name"""
    try:
        list_report_config_data=[]
        df=pd.read_excel(filePath,sheet_name=sheet_name)
        report_names = []
        #gu_names = df.columns
        report_names = df['Report Name']
        return report_names
    except Exception as ex:
        logging.exception(ex)
        return report_names

def get_download_step(filePath,sheet_name):
    """get download step"""  
    try:
        df=pd.read_excel(filePath,sheet_name=sheet_name)
        list_config_data=[]
        for i in df.index.values:
            row_data=df.ix[i,['Id','Name','Type','Description','Remark','referXpath','Class','Value','Value2','Xpath']].to_dict()
            list_config_data.append(row_data)
        #print("data：{0}".format(list_config_data))
        return list_config_data
    except Exception as ex:
        logging.exception(ex)

def run_task(list_config_data,browser):
    """run task"""
    try:
        previous_data_row = None
        for data_row in list_config_data:
            return_value = run_task_by_config(browser,data_row,previous_data_row)
            previous_data_row = data_row
            if not return_value.is_success or not return_value.continiu_run:
                return return_value

        return_data = result(True,'')
        return return_value
    except Exception as ex:
        logging.exception(ex)
        id =  config_data_row['Id']
        return_data = result(False,'task number: '+str(id))
        return return_value


def switch_new_tab(browser,wait_count,index):
    """switch to new tab page"""
    have_new_tab = False
    time.sleep(3)
    try:
        if len(browser.window_handles)>index:
            browser.switch_to_window(browser.window_handles[index])
            have_new_tab =True
            return have_new_tab
        elif wait_count>0:
            wait_count = wait_count-1
            have_new_tab = switch_new_tab(browser,wait_count,index)
            return have_new_tab
        elif wait_count ==0:
            return False
    except Exception as ex:
        logging.exception('switch_new_tab :'+str(ex))
        return False

def get_class_text(driver):
    ''' get text by class '''
    try:
        xpath_country = '/html/body/div[1]/div/div/div[2]/div/div[2]/div/div[4]/div[2]/div/div/ul/li/input'
        driver.find_element_by_xpath(xpath_country).click()
        WebDriverWait(driver,100).until(EC.visibility_of_element_located((By.XPATH,"//div[text()='Zimbabwe']")))
        list_e = driver.find_elements_by_css_selector("[class='select2-result-label ui-select-choices-row-inner']")
        data_list=[]
        for e in list_e:
            value_text = e.text
            temp =[value_text]
            data_list.append(temp)

        excel_helper.create_sheet_name(r'C:\Users\nancy.fu.liu\Desktop\Blackline\blackline\DownloadGtfReportConfig.xlsx','Country Name')
        excel_helper.set_colunm_name(['Country'],r'C:\Users\nancy.fu.liu\Desktop\Blackline\blackline\DownloadGtfReportConfig.xlsx','Country Name')
        excel_helper.write_list_data_to_excel(data_list,r'C:\Users\nancy.fu.liu\Desktop\Blackline\blackline\DownloadGtfReportConfig.xlsx','Country Name')
    except Exception as ex:
        logging.exception('get_class_text :'+str(ex))
        raise

def select_previous_month_lastDay_for_download_report(driver,input_xpath):
    ''' select previous month lastDay for download report '''
    try:
        str_time = get_last_month_shorthand_for_download()

        #input_xpath = '/html/body/div[1]/div/div/div[2]/div/div[2]/div/div[7]/div[2]/div/div[1]/p/input'
        driver.find_element_by_xpath(input_xpath).send_keys(str_time)        
    except Exception as ex:
        logging.exception('select_previous_month_lastDay_for_download_report :'+str(ex))
        raise

def download_report_prevoius_month_lastday(driver,input_xpath):
    try:
        previous_month_lastDay = datetime.date(datetime.date.today().year,datetime.date.today().month,1)-datetime.timedelta(1)
        day_number = str(previous_month_lastDay.day)
        xpath = "//*[text()="+day_number+"]"
        elements = driver.find_elements_by_xpath(xpath)
        if len(elements) == 2:
            elements[1].click()
        elif len(elements) == 1:
            elements[1].click()
        else:
            raise
    except Exception as ex:
        logging.exception('download_report_prevoius_month_lastday :'+str(ex))
        raise

def select_country(driver,select_country_textbox_xpath):
    try:
        if len(PublicData.country_list)>0:
            for country_name in PublicData.country_list:
                on_click(select_country_textbox_xpath,driver)
                country_xpath ="//div[text()='"+country_name["name"]+"']"

                success = wait_elment_click_by_previous_xpath(driver,select_country_textbox_xpath,country_xpath,300)
                if not success:
                    raise
                #on_click(country_xpath,driver,300)
    except Exception as ex:
        logging.exception('select_country :'+str(ex))
        raise

def select_gu(driver,select_gu_textbox_xpath):
    try:
        if len(PublicData.gu_list)>0:
            for gu_name in PublicData.gu_list:
                on_click(select_gu_textbox_xpath,driver)
                gu_xpath ="//div[text()='"+gu_name["name"]+"']"
                #on_click(gu_xpath,driver) 
                success = wait_elment_click_by_previous_xpath(driver,select_gu_textbox_xpath,gu_xpath,300)
                if not success:
                    raise
    except  Exception as ex:
        logging.exception('select_gu :'+str(ex))
        raise

def run_task_by_config(driver,config_data_row,previous_data_row):
    """run task by config file"""
    try:
        action_type =  config_data_row['Type'].strip()
        xpath =  config_data_row['Xpath']
        id =  config_data_row['Id']
        if action_type == 'visit':
            url = config_data_row['Value']
            if len(url) != 0 :
                driver.get(url)
            else:
                return_data = result(False,str(id))
                return return_data
        elif action_type == 'eso_authenticate':
            return_data = eso_authenticate(PublicData.web_driver,PublicData.username,PublicData.password,xpath)
            if not return_data.is_success :
                return return_data
        elif action_type == 'set_file_name':
            PublicData.new_current_download_report_file_path = PublicData.download_report_folder_path+ "\\" + config_data_row['Value'].split('.')[0]+'_'+PublicData.user_name+'.xlsx'
            #PublicData.new_current_download_report_file_path = PublicData.download_report_folder_path+ "\\" + config_data_row['Value'].split('.')[0]+'_'+PublicData.user_name.replace('.','',1)+'.xlsx'
            PublicData.old_current_download_report_file_path = PublicData.download_report_folder_path+ "\\" + config_data_row['Value2']
            #remove_file()
            rename_download_file_name()
            #get_class_text(driver)
        elif action_type == 'click':
            on_click(xpath.strip(),driver)
        elif action_type == 'click_by_text':
            txt = config_data_row['Value']
            type_xpath = config_data_row['Value2']
            target_xpath = "//"+type_xpath+"[text()='"+txt+"']"
            on_click(target_xpath,driver)
        elif action_type == 'wait_previous_click':
            txt = config_data_row['Value']
            type_xpath = config_data_row['Value2']
            target_xpath = "//"+type_xpath+"[text()='"+txt+"']"
            previous_xpath = previous_data_row['Xpath']
            success =  wait_elment_click_by_previous_xpath(driver,previous_xpath,target_xpath,300)
            if not success:
                return_data = result(False,'can\'t find web page element,please check network:'+ str(config_data_row['Id']))
                return return_data
            #on_click(target_xpath,driver)
        elif action_type == 'selectPreviousMonthLastDayForDownloadReport':
            select_previous_month_lastDay_for_download_report(driver,xpath)
        elif action_type ==  'selectGU':
            select_gu(driver,xpath)   
        elif action_type ==  'select_country':
            select_country(driver,xpath)             
        elif action_type == 'clickAndOpenNewTab':
            if PublicData.browser_type == 'chrome':
                have_new_tab = False
                wait_check_time = 0
                while not have_new_tab:
                    on_click(xpath.strip(),driver)
                    time.sleep(3)
                    have_new_tab = switch_new_tab(driver,5,1)  
                    wait_check_time = wait_check_time+18
                    if have_new_tab:
                        break
                    if wait_check_time >300:
                        return_data = result(False,'click genarate button error,please check network',False)
                        return return_data                
        elif  action_type =='saveExcel':
                #pyautogui.typewrite(['down'],interval=0.25)
                #pyautogui.typewrite(['enter'],interval=0.25)
                print(123)
        elif action_type == 'wait':
            #logging.info('wait 5 second')
            #time.sleep(5)
            #logging.info('wait 5 second over')
            #logging.info('wait enter begin')
            #return_data = wait_for_generate(driver)  
            #logging.info('wait enter end')
            #if return_data.is_success:
            #    dom_element_div_accenture =WebDriverWait(driver,600).until(EC.visibility_of_element_located((By.XPATH,xpath.strip())))
            #else:
            #    return return_data
            time.sleep(6)
            #ActionChains(driver).send_keys(Keys.ENTER).perform()
            win32api.keybd_event(13,0,0,0)
            xpath_ui_form = '//*[@id="ui_form"]/span[1]/table/tbody/tr[1]/td/div/table/tbody/tr/td[1]/span/div/a[1]'
            time.sleep(2)
            have_ui_form = element_exist(driver,xpath_ui_form,5)
            if not have_ui_form:
                return_data = result(False,'genarate failed,sign in error',False)
                return return_data 
            dom_element_div_accenture =WebDriverWait(driver,600).until(EC.visibility_of_element_located((By.XPATH,xpath.strip())))
        elif action_type =='check_download_success':
            have_file =  check_rename_download(driver,xpath)
            if not have_file:
                return_data = result(False,'download timeout,step:'+ str(config_data_row['Id']))
                return return_data
        elif action_type =='run_upload_task':
            return_data = run_upload_tasks(driver,PublicData.upload_report_need_prepare_count,PublicData.upload_eale_file_foler,PublicData.upload_apac_na_file_foler)
            return return_data
        elif action_type =='get_prepare_count':    
            prepare_count =WebDriverWait(driver,600).until(EC.presence_of_element_located((By.XPATH,xpath.strip())))
            PublicData.upload_report_need_prepare_count = int(driver.find_element_by_xpath(xpath).text) 
        elif action_type =='check_total_count':    
            WebDriverWait(driver,600).until(EC.presence_of_element_located((By.XPATH,xpath.strip())))
            total_count= int(driver.find_element_by_xpath(xpath).text) 
            if total_count == 0:
                return_data = result(True,'0 item need to do',False)
                return return_data
        elif action_type =='get_approve_count':    
            prepare_count =WebDriverWait(driver,600).until(EC.presence_of_element_located((By.XPATH,xpath.strip())))
            PublicData.approve_report_need_prepare_count = int(driver.find_element_by_xpath(xpath).text) 
        elif action_type =='click_select_last_month_of_upload':
           str_time = get_last_month_shorthand()
           dom_element_time =WebDriverWait(driver,600).until(EC.visibility_of_element_located((By.XPATH,"//a[text()='"+str_time+"']")))
           dom_element_time.click()
        elif action_type =='wait_group_accounts_only':   
            time.sleep(2)
            have_item = element_exist(driver,"//div [@atlas-limitless-row='1']/div[1]/button")
            if not have_item:
                if element_exist(driver,"//div[text()='No results found.']"):
                    return_data = result(True,'0 item need to do',False)
                    return return_data
        elif action_type =='run_approve_task':   
            time.sleep(2)
            return_data = run_approve_tasks(driver,PublicData.approve_report_need_prepare_count)
            return return_data
        else:
            return_data = result(False,'config error,task type:'+action_type)
            return return_data

        return_data = result(True,'')
        return return_data
    except Exception as ex:        
        logging.exception('run_task_by_config:'+str(ex)+'stepId:'+ str(config_data_row['Id']))
        return_data = result(False,'task number: '+str(config_data_row['Id']))
        return return_data
        #return False

def wait_for_generate(driver):
    try:
        xpath_ui_form = '//*[@id="ui_form"]/span[1]/table/tbody/tr[1]/td/div/table/tbody/tr/td[1]/span/div/a[1]'
        have_ui_form = False
        #try:
        #   logging.info('have_ui_form 1:'+str(have_ui_form)) 
        #   have_ui_form = driver.find_element_by_xpath(xpath_ui_form)
        #except :
        #    have_ui_form = False
        #logging.info('have_ui_form 2:'+str(have_ui_form)) 
        count = 0
        while not have_ui_form :
            logging.info('13 1')               

            win32api.keybd_event(13,0,0,0)
            time.sleep(3)
            try:
               have_ui_form = driver.find_element_by_xpath(xpath_ui_form)
            except :
               have_ui_form = False
            logging.info('13 2'+str(have_ui_form))
            count =count+3
            if count >60:
                return_data = result(False,'Generate failed,sign in error',False)
                return return_data   
        dom_element_div_accenture =WebDriverWait(driver,600).until(EC.visibility_of_element_located((By.XPATH,xpath.strip())))
        return_data = result(True,'',True)
        return return_data   
    except Exception as ex:
        logging.info('wait_for_generate')

def check_comment_and_time(driver):
    """check comment and time"""
    try:
        have_comments = element_exist(driver,'//*[@id="ctl00_ctl00_contentBody_cphMain_Comments1_commentsGrid_ctl00__0"]')   
        if not have_comments:
            #add comments
            dom_add_comment_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_Comments1_lbAddComment"]'
            on_click(dom_add_comment_xpath,driver)           
            wait_element(driver,'//*[@id="ctl00_ctl00_contentBody_cphMain_commentDetail_bleCommentDesc"]')
            dom_add_comment_textarea = WebDriverWait(driver,600).until(EC.visibility_of_element_located((By.XPATH,'//*[@id="ctl00_ctl00_contentBody_cphMain_commentDetail_bleCommentDesc"]')))
            dom_add_comment_textarea.send_keys('Confidential - Please contact Cross Border Finance People.Mobility.CBF@accenture.com') 
            save_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_btnSave"]'
            on_click(save_xpath,driver)
            dom_add_comment = WebDriverWait(driver,600).until(EC.visibility_of_element_located((By.XPATH,'//*[@id="ctl00_ctl00_contentBody_cphMain_Comments1_lbAddComment"]')))
                                        
        dom_minutes = WebDriverWait(driver,600).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ctl00_ctl00_contentBody_cphMain_CompletionTimeOptional1_etCompletion_tbMinutes"]')))              
        if dom_minutes.get_attribute('value') !=str(PublicData.upload_time):
            dom_minutes.clear()
            dom_minutes.send_keys(str(PublicData.upload_time))
    except Exception as ex:
        logging.exception('check_comment_and_time: '+str(ex))
        raise

def check_certify_page(driver):
    try:
        dom_element_span_Comments_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_Comments1_lblComments"]'
        WebDriverWait(driver,600).until(EC.visibility_of_element_located((By.XPATH,dom_element_span_Comments_xpath)))
    except Exception as ex:
        logging.exception('check_certify_page: '+str(ex))
        raise

def run_upload_task(driver,upload_eale_folder_path,upload_apac_folder,list_prepare_index,prepare_failed_list,cancel_list):
    """run upload task"""
    try:
        dom_element_span_Comments_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_Comments1_lblComments"]'
        start_actions_xpath = '//*[@id="modules--execGrid-grid"]/div[3]/table/thead/tr/th[1]/span'
        wait_element(driver,start_actions_xpath)
        page_index_xpath ='//*[@id="modules--execGrid-grid"]/div[3]/div[2]/div/input'
        page_number = driver.find_element_by_xpath(page_index_xpath).get_attribute("value")
        index = 1+len(prepare_failed_list) 
        page_index = int(index/50)+1
        if page_number != str(page_index):
            driver.find_element_by_xpath(page_index_xpath).clear()
            driver.find_element_by_xpath(page_index_xpath).send_keys(str(page_index))
        next_btn_edit_xpath = "//div [@atlas-limitless-row='"+str(index)+"']/div[1]/button"
        btn_edit =WebDriverWait(driver,600).until(EC.visibility_of_element_located((By.XPATH,next_btn_edit_xpath)))
        #btn_edit =  driver.find_element_by_xpath("//div [@atlas-limitless-row='"+str(index)+"']/div[1]/button")
        upload_description_text = ''
        target_elment_list = driver.find_elements_by_xpath("//div [@atlas-limitless-row='"+str(index)+"']")
        #unidentified_value = ''
        if len(target_elment_list)>1:
           upload_description_text = target_elment_list[1].find_element_by_css_selector("[class='atlas--ui-grid-cell atlasGrid-cell #modules--execGrid-grid-col-accountName']").text
          
           #unidentified_value = target_elment_list[1].find_element_by_css_selector("[class=''atlas--ui-grid-cell atlasGrid-cell #modules--execGrid-grid-col-amountUnidentified'']").text
        else:
           upload_description_text = target_elment_list[0].find_element_by_css_selector("[class='atlas--ui-grid-cell atlasGrid-cell #modules--execGrid-grid-col-accountName']").text
           #unidentified_value = target_elment_list[0].find_element_by_css_selector("[class=''atlas--ui-grid-cell atlasGrid-cell #modules--execGrid-grid-col-amountUnidentified'']").text
        describtion = upload_description_text
        country_name = get_upload_file_country_name(upload_description_text)
        btn_edit.click()        
        dom_element_span_Comments =WebDriverWait(driver,600).until(EC.presence_of_element_located((By.XPATH,dom_element_span_Comments_xpath)))
        str_balance_value = driver.find_element_by_xpath('//*[@id="ctl00_ctl00_contentBody_cphMain_Unidentified1_tbBalance"]').get_attribute("value")
        
        #cancel_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_SaveCertifyButtons1_btnCancel"]'
        #file_name1 = get_upload_file_name(upload_description_text)       
        #on_click(cancel_xpath,driver)
        #task = {"Country" : country_name,'description' : upload_description_text ,'reason' : '123'}
        #prepare_failed_list.append(task)
        #return

        if str_balance_value.startswith('(') and str_balance_value.endswith(')'):
            balance_number_value  =float('-'+str_balance_value.replace('(','').replace(')','').replace(',',''))
        else:
            balance_number_value  =float( str_balance_value.replace(',',''))

        if balance_number_value == 0:                
            #check comments and time
            check_comment_and_time(driver)
            #click certify
            click_cerfify(driver,cancel_list,prepare_failed_list,list_prepare_index,upload_description_text)   
        else:
            file_name = get_upload_file_name(upload_description_text)
            full_file_name = exist_file(upload_eale_folder_path,upload_apac_folder,file_name)
            if full_file_name != '':
                import_file(driver,upload_eale_folder_path,full_file_name)
                #check redirect page
                dom_element_span_Comments = WebDriverWait(driver,600).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ctl00_ctl00_contentBody_cphMain_Comments1_lblComments"]')))
                after_import_str_balance_value = driver.find_element_by_xpath('//*[@id="ctl00_ctl00_contentBody_cphMain_Unidentified1_tbBalance"]').get_attribute("value")
                if after_import_str_balance_value.startswith('(') and after_import_str_balance_value.endswith(')'):
                    after_import_balance_number_value  =float('-'+after_import_str_balance_value.replace('(','').replace(')','').replace(',',''))
                else:
                    after_import_balance_number_value  =float( after_import_str_balance_value.replace(',',''))

                if after_import_balance_number_value == 0:
                    check_comment_and_time(driver)
                    click_cerfify(driver,cancel_list,prepare_failed_list,list_prepare_index,upload_description_text,True) 
                elif after_import_balance_number_value >= -10 and after_import_balance_number_value <= 10:
                    clear_rounding(driver,after_import_str_balance_value)
                    check_certify_page(driver)
                    after_clear_rounding_str_balance_value = driver.find_element_by_xpath('//*[@id="ctl00_ctl00_contentBody_cphMain_Unidentified1_tbBalance"]').get_attribute("value")
                    if after_clear_rounding_str_balance_value.startswith('(') and after_clear_rounding_str_balance_value.endswith(')'):
                        after_clear_rounding_balance_number_value  =float('-'+after_clear_rounding_str_balance_value.replace('(','').replace(')','').replace(',',''))
                    else:
                        after_clear_rounding_balance_number_value  =float( after_clear_rounding_str_balance_value.replace(',',''))

                    if after_clear_rounding_balance_number_value == 0:
                        check_comment_and_time(driver)
                        click_cerfify(driver,cancel_list,prepare_failed_list,list_prepare_index,upload_description_text,True) 
                    else:
                        after_import_file_cancel_click(driver)
                        task = {"Country" : country_name,'Description' : describtion,'Reason' : 'After clear rounding "Unidentified Difference" != 0'}
                        prepare_failed_list.append(task)
                        cancel_list.append(task)
                elif after_import_balance_number_value < -10 or after_import_balance_number_value > 10:
                    after_import_file_cancel_click(driver)
                    #reconciliations_xpath = '//*[@id="modules--execGrid-records"]'
                    wait_element(driver,start_actions_xpath)
                    task = {"Country" : country_name,'Description' : describtion,'Reason' : 'After import report "Unidentified Difference" >10 or <-10'}
                    prepare_failed_list.append(task)
                    cancel_list.append(task)
            else:
                if balance_number_value >= -10 and balance_number_value <= 10:
                    clear_rounding(driver,str_balance_value)
                    check_certify_page(driver)
                    after_clear_rounding_str_balance_value = driver.find_element_by_xpath('//*[@id="ctl00_ctl00_contentBody_cphMain_Unidentified1_tbBalance"]').get_attribute("value")
                    if after_clear_rounding_str_balance_value.startswith('(') and after_clear_rounding_str_balance_value.endswith(')'):
                        after_clear_rounding_balance_number_value  =float('-'+after_clear_rounding_str_balance_value.replace('(','').replace(')','').replace(',',''))
                    else:
                        after_clear_rounding_balance_number_value  =float( after_clear_rounding_str_balance_value.replace(',',''))

                    if after_clear_rounding_balance_number_value == 0:
                        check_comment_and_time(driver)
                        click_cerfify(driver,cancel_list,prepare_failed_list,list_prepare_index,upload_description_text) 
                    else:
                        #cancel_click(driver)
                        cancel_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_SaveCertifyButtons1_btnCancel"]'
                        on_click(cancel_xpath,driver)
                        task = {"Country" : country_name,'Description' : describtion,'Reason' : 'File missing,after clear rounding  "Unidentified Difference" !=0'}
                        prepare_failed_list.append(task)
                        cancel_list.append(task)
                elif balance_number_value>10 or  balance_number_value<-10:
                    cancel_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_SaveCertifyButtons1_btnCancel"]'
                    on_click(cancel_xpath,driver)
                    #reconciliations_xpath = '//*[@id="modules--execGrid-records"]'
                    wait_element(driver,start_actions_xpath)
                    task = {"Country" : country_name,'Description' : describtion,'Reason' : 'File missing, "Unidentified Difference" >10 or <-10'}
                    prepare_failed_list.append(task)
                    cancel_list.append(task)
    except Exception as ex:
        logging.exception("run_upload_task: "+str(ex))  
        task = {"Country" : country_name,'description' : upload_description_text ,'reason' :"Other, error msg :" + str(ex)}
        prepare_failed_list.append(task)
        init_upload_page(driver)

def init_upload_page(driver):
    try:
        cancel_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_SaveCertifyButtons1_btnCancel"]'
        import_file_page_cancel_xpath = '//*[@id="tbibCancel"]'
        add_comment_page_xpath ='//*[@id="ctl00_ctl00_contentBody_cphMain_btnCancel"]'
        clear_rounding_xpath = '//*[@id="tbibCancel"]'
        if element_exist(driver,clear_rounding_xpath,wait_time = 3):           
            on_click(clear_rounding_xpath,driver)
        if element_exist(driver,add_comment_page_xpath,wait_time = 3):           
            on_click(add_comment_page_xpath,driver)
        if element_exist(driver,import_file_page_cancel_xpath,wait_time = 3):           
            on_click(import_file_page_cancel_xpath,driver)
        if element_exist(driver,cancel_xpath,wait_time = 3):           
            on_click(cancel_xpath,driver)
    except Exception as ex:
        logging.exception('init_upload_page'+str(ex))

def run_approve_tasks(driver,task_count):
    ''' run approve tasks '''
    approve_reject_list = []
    return_data = result(True,'Preview completed',False)        
    try:
        while True:
            is_success = run_approve_task(driver,approve_reject_list,task_count)
            if is_success == '':
                break
        #for task_index in range(0, task_count):
        #    is_success = run_approve_task(driver,approve_reject_list,task_count)
    except Exception as ex:
        logging.exception('run_approve_tasks :' + str(ex))
        return_data.is_success = False
        return_data.msg = 'Preview failed'
    finally:
        if len(approve_reject_list)>0:
            excel_helper.create_sheet_name(PublicData.upload_and_approve_excel_log_file_path,PublicData.log_approve_sheet_name)
            colunm_names = ['Country','Description','Reason']
            excel_helper.set_colunm_name(colunm_names,PublicData.upload_and_approve_excel_log_file_path,PublicData.log_approve_sheet_name)
            for info in approve_reject_list:
                arry_info=[info['Country'],info['Description'],info['Reason']]
                excel_helper.write_row_data(arry_info,PublicData.upload_and_approve_excel_log_file_path,PublicData.log_approve_sheet_name)
        return return_data 

def run_approve_task(driver,approve_reject_list,task_count):
    try:
        dom_element_span_Comments_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_Comments1_lblComments"]'
        start_actions_xpath = '//*[@id="modules--execGrid-grid"]/div[3]/table/thead/tr/th[1]/span'
        wait_element(driver,start_actions_xpath)
        index = 1
        next_btn_cert_xpath = "//div [@atlas-limitless-row='"+str(index)+"']/div[1]/button"
        if not element_exist(driver,next_btn_cert_xpath,60):
            return ''
        btn_cert =WebDriverWait(driver,600).until(EC.visibility_of_element_located((By.XPATH,next_btn_cert_xpath)))        
        approve_description_text = driver.find_element_by_xpath("//div [@atlas-limitless-row='"+str(index)+"']/div[10]").text
        country_name = get_upload_file_country_name(approve_description_text)
        btn_cert.click()

        reject_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_SaveCertifyButtons1_btnReject"]'
        dom_element_span_Comments_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_Comments1_lblComments"]'
        wait_element(driver,dom_element_span_Comments_xpath)
        comment_right = check_comment(driver)
        time_right = check_time(driver,PublicData.approve_time)
        if comment_right and time_right:
            certify_xpath ='//*[@id="ctl00_ctl00_contentBody_cphMain_SaveCertifyButtons1_btnSaveCertify"]'
            on_click(certify_xpath,driver)
            if not element_exist(driver,start_actions_xpath,300):
                back_xpath ='ctl00_ctl00_contentBody_cphMain_SaveCertifyButtons1_btnCancel'
                on_click(back_xpath,driver)
                task_count = task_count+1
                logging.info('approve certify timeout,click back button')
        else:
            check_approve(approve_reject_list,driver,comment_right,time_right,country_name,approve_description_text)
            #tiao zhuan de
            have_new_tab = switch_new_tab(driver,5,1)
            if have_new_tab:
                send_notifications_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_lblTitle"]'
                send_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_btnSend"]'
                wait_element(driver,send_notifications_xpath)
                on_click(send_xpath,driver)
                switch_new_tab(driver,5,0)
        return True
    except Exception as ex:
        approve_task = {"Country" : country_name,'Description' : approve_description_text,"Reason": +'Other, error msg :'+str(ex)}
        approve_reject_list.append(approve_task)
        logging.exception('run_approve_task :' + str(ex))
        return False

def check_approve(approve_reject_list,driver,comment_right,time_right,country_name,upload_description_text):
    try:
        reject_xpath ='//*[@id="ctl00_ctl00_contentBody_cphMain_SaveCertifyButtons1_btnReject"]'
        on_click(reject_xpath,driver)
        select_reason_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_SaveCertifyButtons1_pnlRejectionCodes"]/div[1]'
        wait_element(driver,select_reason_xpath)
        err_xpath = ''
        err_msg = ''
        if  not time_right:
            if not comment_right:
                err_msg = 'No comments and time error'
            else:
                err_msg = 'Time error'
            err_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_SaveCertifyButtons1_ddlDecertificationReasonCode"]/option[10]'
        else:
            err_msg = 'No comments'
            err_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_SaveCertifyButtons1_ddlDecertificationReasonCode"]/option[11]'
        
        on_click(err_xpath,driver)
        time.sleep(1)
        btn_continue_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_SaveCertifyButtons1_btnPSubmit"]'
        on_click(btn_continue_xpath,driver)
        start_actions_xpath = '//*[@id="modules--execGrid-grid"]/div[3]/table/thead/tr/th[1]/span'
        WebDriverWait(driver,300).until(EC.presence_of_element_located((By.XPATH,start_actions_xpath)))

        approve_task = {"Country" : country_name,'description' : upload_description_text,"Reason":err_msg}
        approve_reject_list.append(approve_task)
    except Exception as ex:
        logging.exception('check_approve :'+ str(ex))
        approve_task = {"Country" : country_name,'description' : upload_description_text,"Reason":"Other, error msg :" + str(ex)}
        approve_reject_list.append(approve_task)
        raise


def check_time(driver,right_time):
    ''' check_time '''
    try:
        time_right = False
        time_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_CompletionTimeOptional1_etCompletion_tbMinutes"]'
        time_value =  driver.find_element_by_xpath(time_xpath).get_attribute('value')
        if time_value == str(right_time):
            time_right = True
        return time_right
    except Exception as es:
        logging.exception('check_time :' + str(ex))
        raise

def check_comment(driver):
    ''' check comment '''
    try:
        comment_right = False
        comments_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_Comments1_commentsGrid_ctl00__0"]/td[5]'
        have_comment = element_exist(driver,comments_xpath)
        #if have_comment:
            #defualt_comment_text = 'Confidential - Please contact Cross Border Finance People.Mobility.CBF@accenture.com'
            #comment_text =  driver.find_element_by_xpath(comments_xpath).text
            #if comment_text ==  defualt_comment_text:
            #    comment_right = True
        return have_comment
    except Exception as ex:
        logging.exception('check_have_comment :'+ str(ex))
        return False

def wait_element(driver,xpath,wait_time = 300):
    try:
        WebDriverWait(driver,wait_time).until(EC.presence_of_element_located((By.XPATH,xpath)))
    except Exception as ex:
        logging.exception('wait_element :' +str(xpath))
        raise

def delete_import_file(driver,count = 2):
    try:
        if count>0:
            btn_delete_xpath='//*[@id="ctl00_ctl00_contentBody_cphMain_GLItems1_lnkDelete"]'
            if element_exist(driver,btn_delete_xpath,5):
                on_click(btn_delete_xpath,driver)            
            sub_page_btn_delete_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_btnSave"]'
            wait_element(driver,sub_page_btn_delete_xpath,10)

            select_frist_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_Items1_gvMain_HeaderRow_Parent"]'
            on_click(select_frist_xpath,driver)
            select_second_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_Items2_gvMain_HeaderRow_Parent"]'
            on_click(select_second_xpath,driver)
            on_click(sub_page_btn_delete_xpath,driver)
            wait_element(driver,"//li[text()='Operation Succeeded']")
            sub_page_cancel_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_btnCancel"]'
            on_click(sub_page_cancel_xpath,driver)
            dom_element_span_Comments_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_Comments1_lblComments"]'
            wait_element(driver,dom_element_span_Comments_xpath)
    except Exception as ex:
        logging.exception('delete_import_file,.error msg:'+str(ex) )
        driver.refresh()
        delete_import_file(driver,count-1)

def after_import_file_cancel_click(driver):
    try:
        #delete file
        delete_import_file(driver,2)
        cancel_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_SaveCertifyButtons1_btnCancel"]'
        on_click(cancel_xpath,driver,60)
    except Exception as ex:
        logging.exception('cancel_click:'+str(ex))
        raise

def upload(args):
    pass

def run_upload_tasks(driver,task_count,upload_eale_folder_path,upload_apac_folder):
    prepare_failed_list = []
    cancel_list=[]
    final_prepared_failed_list = []
    frist_time_completed = False
    secend_time_compeleted = False
    return_data = result(True,'Upload completed',False)    
    try:
        if task_count>0:
            for prepare_list_index in range(0, task_count):
                try:
                    run_upload_task(driver,upload_eale_folder_path,upload_apac_folder,prepare_list_index,prepare_failed_list,cancel_list)
                except Exception as ex:
                    return_data.msg = 'Upload failed'
                    return_data.is_success = False
                    error_count = error_count + 1
                    logging.exception("run_upload_task: "+str(ex))  
            frist_time_completed = True
            for upload_failed_item in prepare_failed_list:
                try:
                    row_index = 0
                    run_upload_task(driver,upload_eale_folder_path,upload_apac_folder,row_index,final_prepared_failed_list,cancel_list)
                except Exception as ex:
                    logging.exception("run_upload_task: "+str(ex))  
                    return_data.msg = 'Upload failed'
                    return_data.is_success = False
            secend_time_compeleted = True
    except Exception as ex:
        logging.info('run_upload_tasks:'+str(ex))
        return_data.msg = 'Upload failed'
        return_data.is_success = False
    finally:
        if not frist_time_completed:
            final_prepared_failed_list = prepare_failed_list
        if len(final_prepared_failed_list)>0:
            excel_helper.create_sheet_name(PublicData.upload_and_approve_excel_log_file_path,PublicData.log_upload_sheet_name)
            colunm_names = ['Country','Description','Reason']
            excel_helper.set_colunm_name(colunm_names,PublicData.upload_and_approve_excel_log_file_path,PublicData.log_upload_sheet_name)
            for info in final_prepared_failed_list:
                arry_info=[info['Country'],info['Description'],info['Reason']]
                excel_helper.write_row_data(arry_info,PublicData.upload_and_approve_excel_log_file_path,PublicData.log_upload_sheet_name)  
        return return_data 

def click_cerfify(driver,cancel_list,prepare_failed_list,number,describtion,delete_importfile = False):
    #click certify
    try:
        certify_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_SaveCertifyButtons1_btnSaveCertify"]'
        WebDriverWait(driver,300).until(EC.visibility_of_element_located((By.XPATH,certify_xpath))).click()
        start_actions_xpath = '//*[@id="modules--execGrid-grid"]/div[3]/table/thead/tr/th[1]/span'
        WebDriverWait(driver,300).until(EC.presence_of_element_located((By.XPATH,start_actions_xpath)))
    except Exception as ex:
        logging.exception('click sertify timeout')
        if delete_importfile:
            after_import_file_cancel_click(driver)
        else:
            cancel_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_SaveCertifyButtons1_btnCancel"]'
            on_click(cancel_xpath,driver)
        country_name = get_upload_file_country_name(describtion)
        task = {"Country" : country_name,'Description' : describtion,'Reason' : 'timeout'}
        prepare_failed_list.append(task)
        cancel_list.append(task)



def import_file(driver,upload_folder_path,file_name):
    try: 
        have_import_file = False
        check_have_import_file = '//*[@id="ctl00_ctl00_contentBody_cphMain_GLItems1_lblNoItems"]'
        have_import_file = not element_exist(driver,check_have_import_file)
        have_import_file = False
        if have_import_file:
            btn_delete_xpath='//*[@id="ctl00_ctl00_contentBody_cphMain_GLItems1_lnkDelete"]'
            on_click(btn_delete_xpath,driver)
            sub_page_btn_delete_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_btnSave"]'
            WebDriverWait(driver,600).until(EC.visibility_of_element_located((By.XPATH,sub_page_btn_delete_xpath)))

            select_frist_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_Items1_gvMain_HeaderRow_Parent"]'
            on_click(select_frist_xpath,driver)
            select_second_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_Items2_gvMain_HeaderRow_Parent"]'
            on_click(select_second_xpath,driver)
            on_click(sub_page_btn_delete_xpath,driver)
            WebDriverWait(driver,600).until(EC.presence_of_element_located((By.XPATH,"//li[text()='Operation Succeeded']")))
            sub_page_cancel_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_btnCancel"]'
            on_click(sub_page_cancel_xpath,driver)
        dom_element_span_Comments_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_Comments1_lblComments"]'
        #WebDriverWait(driver,600).until(EC.presence_of_element_located((By.XPATH,dom_element_span_Comments_xpath)))
        wait_element(driver,dom_element_span_Comments_xpath)
        btn_import_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_GLItems1_ibImportItems"]'
        #click import
        on_click(btn_import_xpath,driver)
        #select This month frist day
        import_date_xapth = '//*[@id="ctl00_ctl00_contentBody_cphMain_ftImportSchedule_cdpCloseDate_popupButton"]'
        on_click(import_date_xapth,driver)
        driver.find_element_by_link_text("1").click() 
        list_component_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_ftImportSchedule_ftItemState_RadioButtonL"]'
        driver.find_element_by_xpath(list_component_xpath).click()
        input_file_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_ftSubPreview_oFilefile0"]'
        dom_input_file = WebDriverWait(driver,600).until(EC.presence_of_element_located((By.XPATH,input_file_xpath)))
        dom_input_file.send_keys(upload_folder_path + '\\' + file_name)
        #import sucess
        import_success = find_element_by_class('ruUploadProgress ruUploadSuccess',driver,100)
        if import_success:            
            transactional_amount_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_ftSubPreview_ddlAmt"]'
            driver.find_element_by_xpath(transactional_amount_xpath).click()
            transactional_amount_total_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_ftSubPreview_ddlAmt"]/option[5]'
            on_click(transactional_amount_total_xpath,driver)                 
            account_amount_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_ftSubPreview_ddlAmtLocal"]'
            driver.find_element_by_xpath(account_amount_xpath).click()
            account_amount_total_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_ftSubPreview_ddlAmtLocal"]/option[6]'
            driver.find_element_by_xpath(account_amount_total_xpath).click()
            import_description_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_ftSubPreview_ddlItemDescription"]'
            driver.find_element_by_xpath(import_description_xpath).click()
            description_trsaction_type_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_ftSubPreview_ddlItemDescription"]/option[5]'
            driver.find_element_by_xpath(description_trsaction_type_xpath).click()
            country_name_xpath = '//*[@id="Checkbox1"]'
            on_click(country_name_xpath,driver)
            btn_import_file_xpath = '//*[@id="tbibImport"]'
            driver.find_element_by_xpath(btn_import_file_xpath).click()
        else:
            msg ='Failure:  No records were imported because the data contained the following errors. Please check the data and column mappings and try again.'
            is_err = WebDriverWait(driver,600).until(EC.presence_of_element_located((By.XPATH,"//span[text()='"+msg+"']")))
            if is_err:
                raise Exception("import file failed")
    except Exception as ex:
        logging.exception('import_file :' + str(ex))
        raise

def find_element_by_class(class_name,driver,time=5):
    try:
        have_element = False
        query_time =0
        while query_time<time:
            dom_list = driver.find_elements_by_css_selector("[class='"+class_name+"']")
            if len(dom_list)>0:
                return True
        return have_element
    except Exception as ex:
        logging.exception('find_element_by_class :'+str(ex))
        return False

def clear_rounding(driver,unidentified_difference):
    try:
        add_supporting_item = '//*[@id="ctl00_ctl00_contentBody_cphMain_GLItems1_lbAddItem"]'
        on_click(add_supporting_item,driver)
        amount_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_ftAddItem_tbAmountTxn"]'
        dom_amount = WebDriverWait(driver,600).until(EC.presence_of_element_located((By.XPATH,amount_xpath)))
        dom_amount.send_keys(unidentified_difference)
        time.sleep(1)
        time_clear_rounding_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_ftAddItem_cdpCloseDate_popupButton"]'
        on_click(time_clear_rounding_xpath,driver)
        driver.find_element_by_link_text("1").click() 
        clear_rounding_list_conponent_xpath ='//*[@id="ctl00_ctl00_contentBody_cphMain_ftAddItem_ftItemState_RadioButtonL"]'
        on_click(clear_rounding_list_conponent_xpath,driver)
        text_rounding_xpath = '//*[@id="ctl00_ctl00_contentBody_cphMain_ftAddItem_bleDescription"]'
        driver.find_element_by_xpath(text_rounding_xpath).clear()
        driver.find_element_by_xpath(text_rounding_xpath).send_keys('Roundings')
        save_xpath ='//*[@id="tbibSave"]'
        driver.find_element_by_xpath(save_xpath).click()
    except Exception as ex:
        logging.exception('clear_rounding :'+str(ex))
        raise

def wait_elment_click_by_previous_xpath(driver,previous_xpath,current_xpath,wait_time):
    ''' click element ,check pervious opreation successfully ''' 
    try:
        have_elment = False
        flag_time = 0
        while not have_elment:
            have_elment =  element_exist(driver,current_xpath,50)
            on_click(previous_xpath,driver,10)
            flag_time = flag_time+60
            if flag_time>wait_time:
                break
        if have_elment:
            on_click(current_xpath,driver,10)
            return True
        else:
            return False
    except Exception as ex:
        logging.exception('wait_elment_click_by_previous_xpath:'+current_xpath+' msg:'+str(ex))
        raise

def rename_download_file_name():
    try:
        new_current_download_report_file_path = PublicData.new_current_download_report_file_path
        new_file_name = new_current_download_report_file_path.split('\\')[-1]
        temp_new_file_name = new_file_name.replace('.xlsx','')
        total_count = 0
        if os.path.exists(PublicData.download_report_folder_path):
            for name in os.listdir(PublicData.download_report_folder_path):
                if temp_new_file_name in name:
                    total_count = total_count+1
        if total_count>0:
            for name in os.listdir(PublicData.download_report_folder_path):
                if new_file_name == name:
                    endstr = str(total_count)+'.xlsx'
                    temp_name = new_file_name.replace('.xlsx','')
                    rename_old_file_name = temp_name + endstr
                    os.rename(new_current_download_report_file_path,PublicData.download_report_folder_path+'\\'+rename_old_file_name)
                    break
    except Exception as ex:
        logging.exception('rename_download_file_name :'+str(ex))

def exist_file(eale_file_dir,apac_file_dir,file_name):
    '''
    search file
    '''
    file_full_name =''
    try:
        have_file = False
        if os.path.exists(eale_file_dir) or os.path.exists(apac_file_dir):
            if os.path.exists(eale_file_dir):
                for name in os.listdir(eale_file_dir):
                    if file_name in name:
                        file_full_name = name
                        have_file = True
                        break
            
            if os.path.exists(apac_file_dir) and not have_file:
                for name in os.listdir(apac_file_dir):
                    if file_name in name:
                        file_full_name = name
                        break
        else:
            return ''
        return file_full_name
    except Exception as ex:
        logging.exception('exist_file :' +str(ex))
        return False            

def on_click(xpath,driver,wait_time = 300):
    '''
    click
    '''
    count = 0
    try:
        element = WebDriverWait(driver,wait_time).until(EC.element_to_be_clickable((By.XPATH,xpath)))
        time.sleep(1)
        count = count+1
        element.click()
    except Exception as ex:
        logging.exception('on_click->xpath:'+xpath+'exmsg: '+str(on_click)+' count:'+str(count))
        if count<2:
            element = WebDriverWait(driver,wait_time).until(EC.element_to_be_clickable((By.XPATH,xpath)))
            element.click()
            count = count+1
        logging.exception('on_click->xpath:'+xpath+'exmsg: '+str(on_click)+' count:'+str(count))
        raise

def re_download_report(download_report_list,config_file_path,my_client,count = 2,current_count = 1):
    '''run download task second '''
    download_success = True
    temp__run_task_list = download_report_list
    final_run_task_list = []
    try:
        for item in temp__run_task_list:
            if item['Success'] == True:
                final_info = {'Report name':item['Sheet_name'],'Sheet_name': item['Sheet_name'],'Success':True,'Status': item['Status'],'Info': item['Info']}
                final_run_task_list.append(final_info)
            else:
                final_success = False
                msg = ''
                if current_count == 1:
                    msg = str(current_count)+' time'
                else:
                    msg = str(current_count)+' times'
                my_client.ExecuteJavascript('updateLoadingMsg("Re-downloading '+item['Sheet_name']+' '+msg+'")')

                config_data = get_download_step(config_file_path,item['Sheet_name'])
                for i in range (0,1):
                    PublicData.web_driver = SeleniumHelper.create_driver(PublicData.web_driver,PublicData.currentUserDir,PublicData.myRootPath,PublicData.browser_type,PublicData.download_report_folder_path)
                    final_run_task_result = run_task(config_data,PublicData.web_driver)
                    if final_run_task_result.is_success:
                        final_success = True
                        break
                if final_success:
                    final_info = {"Report name": item['Sheet_name'],'Sheet_name': item['Sheet_name'],'Success':True, 'Status' : 'download successfully','Info' : ''}
                    final_run_task_list.append(final_info)
                else:
                    download_success = False
                    final_info = {"Report name": item['Sheet_name'],'Sheet_name': item['Sheet_name'],'Success':False, 'Status' : 'download failed','Info' : final_run_task_result.msg}
                    final_run_task_list.append(final_info)
        count = count-1
        if count >0:
            current_count = current_count+1
            final_run_task_list = re_download_report(final_run_task_list,config_file_path,my_client,count,current_count)
        
        return final_run_task_list
    except Exception as ex:
        logging.exception('redownload_second :'+str(ex))
        return final_run_task_result

def check_have_avande(list_name):
    is_success = False
    try:
        for name in list_name:
            if 'Ava' in name:
              is_success =  True
        return is_success
    except Exception as ex:
        logging.exception('check_have_avande :'+str(ex))
        return False

def get_upload_file_country_name(discription):
    try:
        have_space_country = ['Costa Rica','Czech Republic','Russian Federation','Saudi Arabia','South Africa','Trinidad,Tobago','United Kingdom']
        list_name = discription.split()
        list_number = discription.split('-')
        if list_name[0]+' '+list_name[1]+' '+list_name[2] == 'United Arab Emirates':
            country_name = 'United Arab Emirates'
        elif discription.startswith('Trinidad and Tobago'):
            country_name = 'Trinidad,Tobago'
        else:
             temp_name=list_name[0]+' '+list_name[1]
             for name in have_space_country:
                if temp_name == name:
                    country_name = temp_name                  
                else:
                    country_name = list_name[0]
        return country_name
    except Exception as  ex:
        logging.exception('get_upload_file_country_name :'+discription)
        return ''

def get_upload_file_name(discription):
    '''get upload file name'''
    file_name = ''
    try:
        special_name = '/141100/236001-236499/271000'
        if special_name in discription:
            return discription

        have_space_country = ['Hong Kong','New Zealand','Puerto Rico','South Korea','Costa Rica','Czech Republic','Russian Federation','Saudi Arabia','South Africa','Trinidad,Tobago','United Kingdom']
        list_name = discription.split()
        list_number = discription.split('-')
        company_name = ''
        country_name =''

        if check_have_avande(list_name):
            company_name = 'Avanade_'
        else :
            company_name = 'Accenture_'

        if list_name[0]+' '+list_name[1]+' '+list_name[2] == 'United Arab Emirates':
            country_name = 'United Arab Emirates'
        elif discription.startswith('Trinidad and Tobago'):
            country_name = 'Trinidad,Tobago'
        else:
             temp_name=list_name[0]+' '+list_name[1]
             for name in have_space_country:
                if temp_name == name:
                    country_name = temp_name  
                    break
                else:
                    country_name = list_name[0]

        file_frist_name = country_name + company_name
        end_number_name = ''
        is_236 = list_number[len(list_number) - 1].strip().startswith('236')
        if is_236:
            end_number_name = '236XXX'
        else :
            end_number_name = list_number[len(list_number) - 1].strip()
        file_name = file_frist_name + end_number_name
        return file_name
    except Exception as ex:
        logging.exception('get_upload_file_name :' + str(ex))
        return file_name
    

def find_upload_file(file_path,file_name):
    if os.path.exists(file_path):
       pass     
        
def element_exist(driver,xpath,wait_time = 60):
    flag=True
    try:
        WebDriverWait(driver,wait_time).until(EC.presence_of_element_located((By.XPATH,xpath)))
        #driver.find_element_by_xpath(xpath)
        return flag        
    except:
        flag=False
        return flag    

def check_download_timeout(browser,xpath_timeout):
    have_window = switch_new_tab(browser,1,2)
    try:
        if have_window:
            dom_element = WebDriverWait(browser,3).until(EC.visibility_of_element_located((By.XPATH,'/html/body/h1')))
            logging.info("check_download_timeout : download file time out")
            return True
    except Exception as ex:        
        return False

def check_rename_download(browser,xpath_timeout):
    try:
        have_file = False
        timeout = False
        wait_time = 0
        is_success = False
        while not have_file:               
            have_file = os.path.exists(PublicData.old_current_download_report_file_path)
            if have_file:
                new_file_name = PublicData.new_current_download_report_file_path
                old_file_name = PublicData.old_current_download_report_file_path
                os.rename(old_file_name,new_file_name)
                have_file = True
                is_success = True
            else:             
                time.sleep(10)
                timeout += 10
                is_time_out = check_download_timeout(browser,xpath_timeout)   
                if is_time_out:
                    is_success = False
                    break
                if timeout>300:
                    logging.info('check_rename_download,timeout>300')
                    is_success = False
                    break
        return  is_success                           
    except Exception as ex:
        logging.exception('check_rename_download :'+str(ex))
        return  False       


def clear_folder(file_path):
    try:
        filelist=os.listdir(file_path)                #列出该目录下的所有文件名
        for f in filelist:
            filepath = os.path.join( file_path, f )   #将文件名映射成绝对路劲
            if os.path.isfile(filepath):            #判断该文件是否为文件或者文件夹
                os.remove(filepath)
        return True
    except Exception as ex:
        logging.exception('clear_folder : '+str(ex))
        return False

def get_download_file_name(config_name):
    try:
        file_name =''
        if config_name.strip() != '':
            config_name.split('.')[0]+PublicData.user_name+'.xlsx'
        return file_name
    except Exception as ex:
        logging.exception('get_download_file_name : '+str(ex))
        return ''

def get_last_month_shorthand():
    '''   get month shortthand  '''
    try:
        last_date_number = datetime.date(datetime.date.today().year,datetime.date.today().month,1)-datetime.timedelta(1)
        today = datetime.date.today()
        first = today.replace(day=1)
        last_month = first - datetime.timedelta(days=1)
        previous_month_number = last_month.month
        month_list =['Jan','Feb','Mar','Apr','May','June','Jul','Aug','Sept','Oct','Nov','Dec']
        str_month_shotthand = month_list[previous_month_number-1] 
        datetime_shotthand = str(last_date_number.day) +' '+str_month_shotthand+' '+str(last_month.year)
        return datetime_shotthand
    except Exception as ex:
        logging.exception("get_last_month_shorthand :" + str(ex))
        raise

def get_last_month_shorthand_for_download():
    '''   get month shortthand  '''
    try:
        last_date_number = datetime.date(datetime.date.today().year,datetime.date.today().month,1)-datetime.timedelta(1)
        today = datetime.date.today()
        first = today.replace(day=1)
        last_month = first - datetime.timedelta(days=1)
        previous_month_number = last_month.month
        month_list =['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
        str_month_shotthand = month_list[previous_month_number-1] 
        datetime_shotthand = str(last_date_number.day) +'-'+str_month_shotthand+'-'+str(last_month.year)
        return datetime_shotthand
    except Exception as ex:
        logging.exception("get_last_month_shorthand :" + str(ex))
        raise

def eso_authenticate_new(driver,username,password,finnal_xpath='',count=5):
    ''' eso authenticate'''
    return_data =result()
    is_success = False
    msg = ''
    try:
        if finnal_xpath != '':
            is_success = element_exist( PublicData.web_driver,finnal_xpath,5)
            if is_success:
                msg = 'Login successfully'
                is_success = True
                return
        if username != '' or password != '':
            if element_exist( PublicData.web_driver,'//*[@id="userNameInput"]',10):
                PublicData.web_driver.find_element_by_xpath('//*[@id="userNameInput"]').send_keys(username)    
                PublicData.web_driver.find_element_by_xpath('//*[@id="passwordInput"]').send_keys(password)  
                PublicData.web_driver.find_element_by_xpath('//*[@id="submitButton"]').click()
                error_text ='Incorrect user ID or password. Type the correct user ID and password, and try again.'
                if element_exist( PublicData.web_driver,"//span[text()='"+error_text+"']",5):
                    msg = "Incorrect user ID or password. Type the correct user ID and password, and try again." 
                    is_success = False
                else:
                    if finnal_xpath != '':
                        is_success = element_exist( PublicData.web_driver,finnal_xpath,60)
                        if is_success:
                            PublicData.username = username
                            PublicData.password = password
                            msg = 'Login successfully'
                        else:
                            is_success = False
                            msg = 'Login failed,type the correct user ID and password, and try again.'
                    else:
                        is_success = True
                        PublicData.username = username
                        PublicData.password = password
                        msg = 'Login successfully'
            else:            
                if count > 0:
                    return_data = eso_authenticate(driver,username,password,finnal_xpath,count - 1)
                    is_success = return_data.is_success
                    msg = return_data.msg
                else:
                    is_success = False
                    msg = 'Login failed,type the correct user ID and password, and try again.'
    except Exception as ex:
        logging.info('eso_authenticate error msg:'+str(ex))
        return_data.is_success = False
        return_data.msg = 'error,please check network,and try again'
    finally :
        return_data.is_success = is_success
        return_data.msg = msg
        return return_data


def eso_authenticate(driver,username,password,finnal_xpath='',count=5):
    ''' eso authenticate'''
    return_data =result()
    is_success = False
    msg = ''
    try:
        if finnal_xpath != '':
            is_success = element_exist( PublicData.web_driver,finnal_xpath,2)
            if is_success:
                PublicData.username = username
                PublicData.password = password
                msg = 'Login successfully'
                is_success = True
                return
        if username != '' or password != '':
            if element_exist( PublicData.web_driver,'//*[@id="userNameInput"]',10):

                page_username = driver.find_element_by_xpath('//*[@id="userNameInput"]').get_attribute('value')

                page_username = driver.find_element_by_xpath('//*[@id="userNameInput"]').get_attribute('value')
                time.sleep(1)
                PublicData.web_driver.find_element_by_xpath('//*[@id="userNameInput"]').clear()
                time.sleep(1)
                PublicData.web_driver.find_element_by_xpath('//*[@id="passwordInput"]').clear()
                time.sleep(1)
                PublicData.web_driver.find_element_by_xpath('//*[@id="userNameInput"]').send_keys(username)   
                time.sleep(1)
                PublicData.web_driver.find_element_by_xpath('//*[@id="passwordInput"]').send_keys(password)  

                PublicData.web_driver.find_element_by_xpath('//*[@id="submitButton"]').click()
                error_text ='Incorrect user ID or password. Type the correct user ID and password, and try again.'
                if element_exist( PublicData.web_driver,"//span[text()='"+error_text+"']",5):
                    msg = "Incorrect user ID or password. Type the correct user ID and password, and try again." 
                    is_success = False
                else:
                    dialog_identity_xpath = '//*[@id="vipDialogIdentityText"]'
                    if element_exist( PublicData.web_driver,dialog_identity_xpath,5):
                        print(123)
                    if finnal_xpath != '':
                        is_success = element_exist( PublicData.web_driver,finnal_xpath,30)
                        if is_success:
                            PublicData.username = username
                            PublicData.password = password
                            msg = 'Login successfully'
                        else:
                            is_success = False
                            msg = 'Login failed,type the correct user ID and password, and try again.'
                    else:
                        is_success = True
                        PublicData.username = username
                        PublicData.password = password
                        msg = 'Login successfully'
            else:            
                if count > 0:
                    return_data = eso_authenticate(driver,username,password,finnal_xpath,count - 1)
                    is_success = return_data.is_success
                    msg = return_data.msg
                else:
                    is_success = False
                    msg = 'Login failed'
    except Exception as ex:
        logging.info('eso_authenticate error msg:'+str(ex))
        return_data.is_success = False
        return_data.msg = 'error,please check network,and try again'
    finally :
        if is_success:
            PublicData.username = username
            PublicData.password = password
            PublicData.user_name = username
        return_data.is_success = is_success
        return_data.msg = msg
        return return_data

def read_excel_by_column_names(filePath,sheet_name,array_column_names):
    """get report list"""
    try:
        df=pd.read_excel(filePath,sheet_name=sheet_name)
        list_report_config_data=[]
        for i in df.index.values:
            row_data=df.ix[i,array_column_names].to_dict()
            list_report_config_data.append(row_data)      
        return list_report_config_data
    except Exception as ex:
        logging.exception(ex)

def get_region_name(gu_list,country_list,gu_country_region_list):
    ''' get region name '''
    msg =''
    is_success = True
    region_name = ''
    try:
        if len(country_list) > 0 and len(gu_list) == 0:
            for country_name_item in country_list:
                country_name =country_name_item['name'].strip()
                for region_config in gu_country_region_list:
                    if region_config[PublicData.region_config_column_name_country_name].strip() == country_name :
                        if region_name != '' and region_name !=  region_config[PublicData.region_config_column_name_region].strip() :
                            is_success = False
                            msg ='You selected country within two regions'
                            return
                        region_name =  region_config[PublicData.region_config_column_name_region] 
        elif len(country_list) > 0 and  len(gu_list)>0:
            for country_name_item in country_list:
                country_name =country_name_item['name'].strip()
                for region_config in gu_country_region_list:
                    if region_config[PublicData.region_config_column_name_country_name].strip() == country_name:
                        gu_name = region_config[PublicData.region_config_column_name_gu]
                        country_has_mapping_gu = False
                        for selected_gu_name in gu_list:
                            if selected_gu_name['name'].strip() == gu_name:
                                country_has_mapping_gu = True
                                break
                        if not country_has_mapping_gu:
                            is_success = False
                            msg = 'Gu and country dismatch '
                            return
            for gu_name_item in gu_list:
                gu_name = gu_name_item['name'].strip()
                for region_config in gu_country_region_list:
                    if region_config[PublicData.region_config_column_name_gu].strip() == gu_name :
                        if region_name != '' and region_name !=  region_config[PublicData.region_config_column_name_region].strip() :
                            is_success = False
                            msg ='You selected gu within two regions'
                            logging.info(msg+':'+gu_name)
                            return
                        region_name =  region_config[PublicData.region_config_column_name_region].strip()
        elif len(country_list) == 0 and len(gu_list) > 0:
            for gu_name_item in gu_list:
                gu_name = gu_name_item['name'].strip()
                for region_config in gu_country_region_list:
                    if region_config[PublicData.region_config_column_name_gu].strip() == gu_name :
                        if region_name != '' and region_name !=  region_config[PublicData.region_config_column_name_region].strip() :
                            is_success = False
                            msg ='You selected gu within two regions'
                            return
                        region_name =  region_config[PublicData.region_config_column_name_region] 
    except Eception as ex:
        logging.exception('get_region_name :'+str(ex))
    finally:
        result_data =  {"is_success": is_success,'msg': msg,'region_name':region_name}
        return result_data

