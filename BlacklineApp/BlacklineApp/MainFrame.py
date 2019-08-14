import json
import wx
from cefpython3 import cefpython as cef
import platform
import sys
import os
import threading
import Utility
import PublicData
import SeleniumHelper
from return_value_model import return_data as result
from return_value_model import return_font_data
import logging
import log_helper
import excel_helper
import call_vba
import time
#import urllib


# Platforms
WINDOWS = (platform.system() == "Windows")
LINUX = (platform.system() == "Linux")
MAC = (platform.system() == "Darwin")

if MAC:
    try:
        # noinspection PyUnresolvedReferences
        from AppKit import NSApp
    except ImportError:
        print("[wxpython.py] Error: PyObjC package is missing, "
              "cannot fix Issue #371")
        print("[wxpython.py] To install PyObjC type: "
              "pip install -U pyobjc")
        sys.exit(1)

# Configuration
#WIDTH = 1100
#HEIGHT = 700

WIDTH = 800
HEIGHT = 500

# Globals
g_count_windows = 0

def re_set_window(width,height):
    try:
        size = scale_window_size_for_high_dpi(width, height)
        PublicData.MainFrom.SetSize(size)
        PublicData.MainFrom.Centre()
        #PublicData.MainFrom.SetPosition()
        #wx.Frame.__init__(self, parent=None, id=wx.ID_ANY,
        #            title='Blackline Autometion', size=size,pos=(110,10))
    except Exception as ex:
        logging('re_set_window')

def scale_window_size_for_high_dpi(width, height):
    """Scale window size for high DPI devices. This func can be
    called on all operating systems, but scales only for Windows.
    If scaled value is bigger than the work area on the display
    then it will be reduced."""
    if not WINDOWS:
        return width, height
    (_, _, max_width, max_height) = wx.GetClientDisplayRect().Get()
    # noinspection PyUnresolvedReferences
    (width, height) = cef.DpiAware.Scale((width, height))
    if width > max_width:
        width = max_width
    if height > max_height:
        height = max_height
    return width, height

class MainFrame(wx.Frame):

    def __init__(self):
        self.browser = None
        # Must ignore X11 errors like 'BadWindow' and others by
        # installing X11 error handlers. This must be done after
        # wx was intialized.
        if LINUX:
            cef.WindowUtils.InstallX11ErrorHandlers()

        global g_count_windows
        g_count_windows += 1

        if WINDOWS:
            # noinspection PyUnresolvedReferences, PyArgumentList
            print("[wxpython.py] System DPI settings: %s"
                  % str(cef.DpiAware.GetSystemDpi()))
        if hasattr(wx, "GetDisplayPPI"):
            print("[wxpython.py] wx.GetDisplayPPI = %s" % wx.GetDisplayPPI())
        print("[wxpython.py] wx.GetDisplaySize = %s" % wx.GetDisplaySize())

        print("[wxpython.py] MainFrame declared size: %s"
              % str((WIDTH, HEIGHT)))
        size = scale_window_size_for_high_dpi(WIDTH, HEIGHT)
        print("[wxpython.py] MainFrame DPI scaled size: %s" % str(size))

        wx.Frame.__init__(self, parent=None, id=wx.ID_ANY,
                          title='Blackline Autometion', size=size)
        # wxPython will set a smaller size when it is bigger
        # than desktop size.
        print("[wxpython.py] MainFrame actual size: %s" % self.GetSize())

        iconPath = PublicData.myRootPath+r'\Page\img\favicon.ico'
        self.icon1 = wx.Icon(name=iconPath,  type=wx.BITMAP_TYPE_ICO)
        self.SetIcon(self.icon1)

        #self.setup_icon()
        #self.create_menu()
        self.Bind(wx.EVT_CLOSE, self.OnClose)

        # Set wx.WANTS_CHARS style for the keyboard to work.
        # This style also needs to be set for all parent controls.
        self.browser_panel = wx.Panel(self, style=wx.WANTS_CHARS)
        self.browser_panel.Bind(wx.EVT_SET_FOCUS, self.OnSetFocus)
        self.browser_panel.Bind(wx.EVT_SIZE, self.OnSize)

        if MAC:
            # Make the content view for the window have a layer.
            # This will make all sub-views have layers. This is
            # necessary to ensure correct layer ordering of all
            # child views and their layers. This fixes Window
            # glitchiness during initial loading on Mac (Issue #371).
            NSApp.windows()[0].contentView().setWantsLayer_(True)

        if LINUX:
            # On Linux must show before embedding browser, so that handle
            # is available (Issue #347).
            self.Show()
            # In wxPython 3.0 and wxPython 4.0 on Linux handle is
            # still not yet available, so must delay embedding browser
            # (Issue #349).
            if wx.version().startswith("3.") or wx.version().startswith("4."):
                wx.CallLater(100, self.embed_browser)
            else:
                # This works fine in wxPython 2.8 on Linux
                self.embed_browser()
        else:
            self.embed_browser()
            self.Show()

    def embed_browser(self):
        window_info = cef.WindowInfo()
        (width, height) = self.browser_panel.GetClientSize().Get()
        assert self.browser_panel.GetHandle(), "Window handle not available"
        window_info.SetAsChild(self.browser_panel.GetHandle(),
                               [0, 0, width, height])
        self.browser = cef.CreateBrowserSync(window_info,
                                             url="file:///Page//login.html")
        self.browser.SetClientHandler(FocusHandler())

        js = cef.JavascriptBindings( bindToFrames=True, bindToPopups=False)
        self.bind_function(js)
        
        self.browser.SetJavascriptBindings(js)
        #self.browser.HasDevTools()
    
    def bind_function(self,js):
        ''' bind function for font page ''' 
        try:
            js.SetFunction("downloadReport",self.download_gtf_report)
            js.SetFunction("getGuNamesFrm",self.get_gu_name)
            js.SetFunction("runVbaFrm",self.run_vba)
            js.SetFunction("uploadFrm",self.upload)
            js.SetFunction("previewFrm",self.preview)
            js.SetFunction("runTasksFrm",self.run_automation_tasks)
            js.SetFunction("loadingFrm",self.on_loading_page)
            js.SetFunction("closeDevToolsFrm",self.closeDevToolsFrm)
            js.SetFunction("showDevToolsFrm",self.showDevToolsFrm)
            js.SetFunction("checkUserInfo",self.check_user_info)
        except Exception as ex:
            logging.exception('bind_function :'+str(ex))
    
    def check_user_info(self,username,password):
        try:
            my_client = self.browser.GetFrames()[0]
            thread_run_tasks=threading.Thread(target=self.check_gtf_permission,args=(my_client,username,password))
            thread_run_tasks.setDaemon(True)
            thread_run_tasks.start()
        except Exception as ex:
            logging.exception('check_user_info: '+str(ex))

    def check_gtf_permission(self,my_client,username,password):
        msg = ''
        is_success = False
        return_data = result()
        try:
            my_client.ExecuteJavascript('showLoading()') 
            if username != '' or password != '':
                PublicData.web_driver = SeleniumHelper.create_driver(PublicData.web_driver,PublicData.currentUserDir,PublicData.myRootPath,'chrome',PublicData.download_report_folder_path,True,True)
                if not PublicData.web_driver:
                    msg = "create browser failed ,please close all chrome browser"
                    return_data.is_success = False
                    return_data.msg = msg
                    return
                PublicData.web_driver.get('https://gtf.accenture.com/reports')
                xpath = '//*[@id="content"]/div[1]/div[1]/div[1]/div[2]/div[2]/div[11]/div/div[3]/button'
                return_data = Utility.eso_authenticate(PublicData.web_driver,username,password,xpath)
        except Exception as ex:
            logging.info('check_user_info:'+str(ex))
            return_data.is_success = False
            return_data.msg = 'Authentication failed,please check network and try again'
        finally:
            my_client.ExecuteJavascript('hideLoading()') 
            if return_data.is_success:
                my_client.ExecuteJavascript('redirectiUrl("index.html")') 
                re_set_window(1100,700)
            else:
                if return_data.msg == '':
                    return_data.msg = 'Authentication failed,please check network and try again'
                my_client.ExecuteJavascript('showMessage("'+return_data.msg+'")')
                return_data.msg = ''
            SeleniumHelper.close_driver(PublicData.web_driver)

    def showDevToolsFrm(self):
        try:
           self.browser.ShowDevTools()  
           print(123)
        except Exception as ex:
            print(ex)

    def set_region_refer_path(self):
        try:
            PublicData.download_report_folder_path_total =  r'C:\Blackline Automation App\Report'+'\\'+PublicData.app_previous_month_folder_name+'\\'+PublicData.download_report_region_name+r'\Source File'
            PublicData.blackline_monthly_reconciliation_previous_month_file_path =r'C:\Blackline Automation App\Report'+'\\'+PublicData.app_previous_month_folder_name+'\\'+PublicData.download_report_region_name+r'\BlacklineMonthlyReconciliation\\'+PublicData.report_name_in_monthly_reconciliation_folder
            PublicData.upload_report_folder = r'C:\Blackline Automation App\Report'+'\\'+PublicData.app_previous_month_folder_name+'\\'+PublicData.download_report_region_name+r'\ImportGroup'
            PublicData.download_report_folder_path =  r'C:\Blackline Automation App\Report'+'\\'+PublicData.app_previous_month_folder_name+'\\'+PublicData.download_report_region_name+r'\Source File_'+ PublicData.username
            PublicData.upload_and_approve_excel_log_file_path = r'C:\Blackline Automation App\Report'+'\\'+PublicData.app_previous_month_folder_name+'\\'+PublicData.download_report_region_name+r'\BlacklineMonthlyReconciliation\Report_May19.xlsx'
            PublicData.generate_report_blackline_folder_path_by_vba_tool =r'C:\Blackline Automation App\Report'+'\\'+PublicData.app_previous_month_folder_name+'\\'+PublicData.download_report_region_name+r'\BlacklineMonthlyReconciliation'
            PublicData.generate_report_import_folder_path_by_vba_tool =r'C:\Blackline Automation App\Report'+'\\'+PublicData.app_previous_month_folder_name+'\\'+PublicData.download_report_region_name+r'\ImportGroup'
            PublicData.vba_tool_folder = r'C:\Blackline Automation App\Report'+'\\'+PublicData.app_previous_month_folder_name+'\\'+PublicData.download_report_region_name
            PublicData.template_folder = r'C:\Blackline Automation App\Report'+'\\'+PublicData.app_previous_month_folder_name+'\\'+PublicData.download_report_region_name+r'\Template'
            PublicData.template_MonthlyReconReportTemplate_file_path = r'C:\Blackline Automation App\Report'+'\\'+PublicData.app_previous_month_folder_name+'\\'+PublicData.download_report_region_name+r'\Template\Monthly Recon Report_Template.xlsx'
            PublicData.template_ReportTemplate_file_path = r'C:\Blackline Automation App\Report'+'\\'+PublicData.app_previous_month_folder_name+'\\'+PublicData.download_report_region_name+r'\Template\Report_Template.xlsx'
        except Exception as ex:
            logging('set_region_refer_path :'+str(ex))

    def closeDevToolsFrm(self):
        try:
          data =  self.browser.CloseDevTools()
          print(123)
        except Exception as ex:
            print(ex)    

    def run_automation_tasks(self,browser_type,str_gu_list,str_country_list,time):
        '''create download upload vba task thread '''
        msg = ''
        try:
            PublicData.gu_list = json.loads(str_gu_list)
            PublicData.country_list = json.loads(str_country_list)

            column_names = [PublicData.region_config_column_name_country_name,PublicData.region_config_column_name_gu,PublicData.region_config_column_name_region]
            PublicData.region_config_list =  Utility.read_excel_by_column_names(PublicData.task_config_file_path,PublicData.region_config_sheet_name,column_names)
            return_result = Utility.get_region_name(PublicData.gu_list,PublicData.country_list,PublicData.region_config_list)
            if return_result['is_success'] == True:
                PublicData.download_report_region_name =  return_result['region_name']
               # PublicData.set_region_refer_path()
                self.set_region_refer_path()
            else:
                self.browser.GetFrames()[0].ExecuteJavascript('hideLoading()')
                msg = return_result['msg']   
                self.browser.GetFrames()[0].ExecuteJavascript('showMessage("'+msg+'")')  
                return

            have_file = Utility.init_app_need_file()   
            if not have_file:
                self.browser.GetFrames()[0].ExecuteJavascript('hideLoading()')
                self.browser.GetFrames()[0].ExecuteJavascript('showMessage("init file error")')
                return

            my_client = self.browser.GetFrames()[0]
            thread_run_tasks=threading.Thread(target=self.run_download_vba_upload_tasks,args=(my_client,browser_type,str_gu_list,str_country_list,time))
            thread_run_tasks.setDaemon(True)
            thread_run_tasks.start()
        except Exception as ex:
            logging.exception('run_automation_tasks: '+str(ex))
            msg = 'Run generate report task failed'
            self.browser.GetFrames()[0].ExecuteJavascript('hideLoading()')
            self.browser.GetFrames()[0].ExecuteJavascript('showMessage("'+msg+'")')   

    def run_download_vba_upload_tasks(self,my_client,browser_type,str_gu_list,str_country_list,upload_time):
        '''run download upload vba task '''
        logging.info('run_download_vba_upload_tasks begin')
        is_success = False
        msg = ''
        try:
            config_file_path = PublicData.task_config_file_path
            return_value = self.run_download_report_tasks(config_file_path,browser_type,my_client,1,'',False)
            if not return_value.is_success:
                my_client.ExecuteJavascript('showMessage("'+return_value.msg+'","'+return_value.title+'","'+str(return_value.width)+'")')
                return
            return_value = self.run_vba_task(my_client,False)
            if not return_value.is_success:
                my_client.ExecuteJavascript('showMessage("'+return_value.msg+'")')
                return
            return_value = self.run_upload_task(my_client,'chrome',False,upload_time)
            if not return_value.is_success:
                my_client.ExecuteJavascript('showMessage("'+return_value.msg+'")')
                return
            else:
                my_client.ExecuteJavascript('showMessage("Generate report completed")')
            logging.info('run_download_vba_upload_tasks end')
        except Exception as ex:
            logging.exception('run_download_vba_upload_tasks: '+str(ex))
            my_client.ExecuteJavascript('hideLoading()')
            my_client.ExecuteJavascript('showMessage("Generat report failed")')
            SeleniumHelper.close_driver(PublicData.web_driver)
        finally:
            my_client.ExecuteJavascript('hideLoading()')
            Utility.copy_log_info(PublicData.download_report_excel_sheet_name)
            Utility.copy_log_info(PublicData.log_upload_sheet_name)
            Utility.copy_log_info(PublicData.log_approve_sheet_name)

    def preview(self,browser_type,approve_time):
        '''create preview thread '''
        try:
            PublicData.approve_time = int(approve_time)
            my_client = self.browser.GetFrames()[0]
            thread_preview=threading.Thread(target=self.run_preview_task,args=(my_client,browser_type))
            thread_preview.setDaemon(True)
            thread_preview.start()
        except Exception as ex:
            logging.exception('preview :'+str(ex))
            self.browser.GetFrames()[0].ExecuteJavascript('hideLoading()')
            self.browser.GetFrames()[0].ExecuteJavascript('showMessage("Preview failed")')
    
    def on_loading_page(self):
        ''' font page init '''
        try:
            column_names = [PublicData.region_config_column_name_country_name,PublicData.region_config_column_name_gu,PublicData.region_config_column_name_region]
            PublicData.region_config_list =  Utility.read_excel_by_column_names(PublicData.task_config_file_path,PublicData.region_config_sheet_name,column_names)

            #have_file_log_file = Utility.create_file(PublicData.app_log_file_path)
            #log_helper.config_log_info(PublicData.app_log_file_path) 

            #have_file = Utility.init_app_need_file()   
            #if not have_file:
            #    self.browser.GetFrames()[0].ExecuteJavascript('showMessage("init file error")')
            logging.info('on_loading_page end')
        except Exception as ex:
            logging.exception('log_helper on_loading_page error ,msg:'+str(ex))
            
    def run_preview_task(self,client,browser_type):
        ''' run preview task'''
        try:
            logging.info('run_preview_task begin')
            config_data = Utility.get_download_step(PublicData.task_config_file_path,PublicData.approve_task_config_sheet_name)
            PublicData.web_driver = SeleniumHelper.create_driver(PublicData.web_driver,PublicData.currentUserDir,PublicData.myRootPath,browser_type,PublicData.download_report_folder_path)
            if not PublicData.web_driver:
                msg = "create browser failed ,please close all chrome browser"
                return  
            client.ExecuteJavascript('updateLoadingMsg("Runing preview")')
            run_task_result =  Utility.run_task(config_data,PublicData.web_driver)
            if run_task_result.is_success:
                msg = run_task_result.msg
            else:
                msg = 'Preview failed'
            logging.info('run_preview_task end')
        except Exception as ex:
            logging.exception('run_preview_task: '+str(ex))
            msg = 'Preview failed'
        finally:
            client.ExecuteJavascript('hideLoading()')
            client.ExecuteJavascript('showMessage("'+msg+'")')
            SeleniumHelper.close_driver(PublicData.web_driver)
            Utility.copy_log_info(PublicData.download_report_excel_sheet_name)
            Utility.copy_log_info(PublicData.log_upload_sheet_name)
            Utility.copy_log_info(PublicData.log_approve_sheet_name)

    def upload(self,browser_type,input_time):
        '''create upload thread '''
        try:
            my_client = self.browser.GetFrames()[0]
            thread_upload=threading.Thread(target=self.run_upload_task,args=(my_client,browser_type,True,input_time))
            thread_upload.setDaemon(True)
            thread_upload.start()
        except Exception as ex:
            logging.exception('upload :'+str(ex))
            my_client.ExecuteJavascript('hideLoading()')
            my_client.ExecuteJavascript('showMessage("Upload failed")')
    
    def run_upload_task(self,client,browser_type,alert_msg = True,upload_time =5):
        ''' run upload task '''
        return_value =  return_font_data()
        PublicData.upload_time = int(upload_time)
        msg = ''
        is_success = False
        try:
            logging.info('run_upload_task begin')
            config_data = Utility.get_download_step(PublicData.task_config_file_path,PublicData.upload_task_config_sheet_name)
            PublicData.web_driver = SeleniumHelper.create_driver(PublicData.web_driver,PublicData.currentUserDir,PublicData.myRootPath,browser_type,PublicData.download_report_folder_path)
            if not PublicData.web_driver:
                msg = "Create browser failed ,please close all chrome browser"
                return  
            client.ExecuteJavascript('updateLoadingMsg("Running upload")')
            run_task_result =  Utility.run_task(config_data,PublicData.web_driver)
            if run_task_result.is_success:
                msg = run_task_result.msg
                is_success = True
            else:
                msg = 'Upload failed'
            logging.info('run_upload_task end')
        except Exception as ex:
            logging.exception('run_upload_task: '+str(ex))
            msg = 'Upload failed'
        finally:
            client.ExecuteJavascript('hideLoading()')
            if alert_msg:
                client.ExecuteJavascript('showMessage("'+msg+'")')   
            
            return_value.is_success =is_success
            return_value.msg = msg
            SeleniumHelper.close_driver(PublicData.web_driver)
            Utility.copy_log_info(PublicData.download_report_excel_sheet_name)
            Utility.copy_log_info(PublicData.log_upload_sheet_name)
            Utility.copy_log_info(PublicData.log_approve_sheet_name)
            return return_value


    def run_vba(self,str_gu_list,str_country_list):
        '''create run vba  thread '''
        try:
            PublicData.gu_list = json.loads(str_gu_list)
            PublicData.country_list = json.loads(str_country_list)

            column_names = [PublicData.region_config_column_name_country_name,PublicData.region_config_column_name_gu,PublicData.region_config_column_name_region]
            PublicData.region_config_list =  Utility.read_excel_by_column_names(PublicData.task_config_file_path,PublicData.region_config_sheet_name,column_names)
            return_result = Utility.get_region_name(PublicData.gu_list,PublicData.country_list,PublicData.region_config_list)
            if return_result['is_success'] == True:
                PublicData.download_report_region_name =  return_result['region_name']
                #PublicData.set_region_refer_path()
                self.set_region_refer_path()
            else:
                self.browser.GetFrames()[0].ExecuteJavascript('hideLoading()')
                msg = return_result['msg']   
                self.browser.GetFrames()[0].ExecuteJavascript('showMessage("'+msg+'")')  
                return

            have_file = Utility.init_app_need_file()   
            if not have_file:
                self.browser.GetFrames()[0].ExecuteJavascript('hideLoading()')
                self.browser.GetFrames()[0].ExecuteJavascript('showMessage("init file error")')
                return            

            my_client = self.browser.GetFrames()[0]
            thread_run_vba=threading.Thread(target=self.run_vba_task,args=(my_client,True))
            thread_run_vba.setDaemon(True)
            thread_run_vba.start()
        except Exception as ex:
            logging.exception('run_vba :'+str(ex))
            self.browser.GetFrames()[0].ExecuteJavascript('hideLoading()')
            self.browser.GetFrames()[0].ExecuteJavascript('showMessage("Run vba failed")')

    def run_vba_task(self,client,alert_msg = True):
        return_value =  return_font_data()
        is_success = False
        msg = ''
        try:  
            logging.info('run_vba_task begin')            
            have_file = Utility.check_vba_tool_need_files()
            if have_file:
                tool_path = PublicData.vba_tool_folder+'\\tool.xlsm'
                #run vba
                client.ExecuteJavascript('updateLoadingMsg("Running import report function")')
                import_result = call_vba.run_vba_blackline_macro(tool_path,PublicData.vba_tool_import_macro_name)
                time.sleep(2)
                client.ExecuteJavascript('updateLoadingMsg("Running generate function")')
                generate_result = call_vba.run_vba_blackline_macro(tool_path,PublicData.vba_tool_generate_macro_name)
                if import_result and generate_result:
                    is_success = True
                    msg = 'Run vba successfully'
                elif not import_result and not generate_result:
                    msg = 'Run vba import function and generate function failed'
                elif not import_result:
                    msg = 'Run vba import function failed'
                elif not generate_result:
                    msg = 'Run vba generate function failed'
            else:
                client.ExecuteJavascript('hideLoading()')
                msg = "Run vba failed,there is no source files"
                is_success = False
            logging.info('run_vba_task end')  
        except Exception as ex:
            logging.exception('run_vba :'+str(ex))
            msg = 'Run vba failed'
        finally :
            if alert_msg:
                client.ExecuteJavascript('hideLoading()')
                client.ExecuteJavascript('showMessage("'+msg+'")')            
            return_value.is_success = is_success
            return_value.msg = msg
            return return_value

    def get_gu_name(self):
        try:
            #have_file = Utility.init_app_need_file()
            have_file = True
            if have_file:
                report_names = Utility.get_download_report_names(PublicData.task_config_file_path,PublicData.config_sheet_name)
                if len(report_names) == 0:
                   self.browser.GetFrames()[0].ExecuteJavascript('showMessage("get gu names failed")')

                dict_report_names = []
                for gu_name in report_names:
                     name = self.create_report_name_dict(gu_name)
                     dict_report_names.append(name)
                jsondata = json.dumps(dict_report_names)
                self.browser.GetFrames()[0].ExecuteJavascript('bindReportNames('+jsondata+')')
            else:
                self.browser.GetFrames()[0].ExecuteJavascript('showMessage("Create file failed,please check file path")')
        except Exception as ex:
            logging.exception('get_gu_name :'+ str(ex))
        finally:
            self.browser.GetFrames()[0].ExecuteJavascript('hideLoading()')

    def create_report_name_dict(self,name):
        return {
                'name': name
                }
         
    def download_gtf_report(self,browser_type,str_gu_list,str_country_list,download_type=1,report_name=''):
        ''' download reprt
            para :
            download_type:1  whole   2:steps
        '''
        msg = ''
        try:
            download_type = int(download_type)
            PublicData.gu_list = json.loads(str_gu_list)
            PublicData.country_list = json.loads(str_country_list)

            column_names = [PublicData.region_config_column_name_country_name,PublicData.region_config_column_name_gu,PublicData.region_config_column_name_region]
            PublicData.region_config_list =  Utility.read_excel_by_column_names(PublicData.task_config_file_path,PublicData.region_config_sheet_name,column_names)
            return_result = Utility.get_region_name(PublicData.gu_list,PublicData.country_list,PublicData.region_config_list)
            if return_result['is_success'] == True:
                PublicData.download_report_region_name =  return_result['region_name']
                #PublicData.set_region_refer_path()
                self.set_region_refer_path()
            else:
                self.browser.GetFrames()[0].ExecuteJavascript('hideLoading()')
                msg = return_result['msg']   
                self.browser.GetFrames()[0].ExecuteJavascript('showMessage("'+msg+'")')  
                return

            have_file = Utility.init_app_need_file()   
            if not have_file:
                self.browser.GetFrames()[0].ExecuteJavascript('hideLoading()')
                self.browser.GetFrames()[0].ExecuteJavascript('showMessage("init file error")')
                return

            PublicData.browser_type = browser_type
            config_file_path = PublicData.task_config_file_path

            my_client = self.browser.GetFrames()[0]
            thread_download=threading.Thread(target=self.run_download_report_tasks,args=(config_file_path,browser_type,my_client,download_type,report_name,True))
            thread_download.setDaemon(True)
            thread_download.start()
        except Exception as ex:    
            msg = 'Download Failed'
            logging.exception('download_gtf_report :'+ str(ex))
            self.browser.GetFrames()[0].ExecuteJavascript('hideLoading()')
            self.browser.GetFrames()[0].ExecuteJavascript('showMessage("'+msg+'")')    

    def run_download_report_tasks(self,config_file_path,browser_type,my_client,download_type,report_name,alert_msg = True):
        '''run download report tasks '''
        return_value =  return_font_data()
        msg = ''
        is_success = True
        titile = 'Download report'
        final_run_task_list = []
        try: 
            logging.info('run_vba_task begin')  
            clean_success = False
            #if download_type == 1:
            #    clean_success = Utility.clear_folder(PublicData.download_report_folder_path)
            #read config
            list_config_report_data = self.get_task_config_data(config_file_path,download_type,report_name)

            if list_config_report_data != None and len(list_config_report_data)>0:
                download_report_list = []
                for data_row in list_config_report_data:  
                    time.sleep(1)
                    max_download_count = int(data_row[PublicData.str_excel_config_colunm_redownload_count])
                    sheet_name = data_row[PublicData.str_excel_config_colunm_report_name]
                    #self.set_download_file_name(data_row)    

                    #if download_type == 2:
                    #    self.remove_file()
                            
                    config_data = Utility.get_download_step(config_file_path,sheet_name)
                    my_client.ExecuteJavascript('updateLoadingMsg("Downloading '+sheet_name+'")')
                    if config_data == None:
                        info = {"Report name": sheet_name,'Success':False,'Status' : 'download failed','Info' : 'sheet name error','Sheet_name':sheet_name}
                        download_report_list.append(info)
                        continue

                    for i in range (0,max_download_count):
                        logging.info("download report start: "+sheet_name+',count:'+str(i+1))
                        PublicData.web_driver = SeleniumHelper.create_driver(PublicData.web_driver,PublicData.currentUserDir,PublicData.myRootPath,browser_type,PublicData.download_report_folder_path)

                        if PublicData.web_driver == None:
                            return_value.msg = "Create browser failed ,please close all chrome browser"                            
                            return_value.is_success = False
                            return_value.width = 400
                            return  
                        
                        run_task_result =  Utility.run_task(config_data,PublicData.web_driver)
                        if run_task_result.is_success:
                            break

                    if not run_task_result.is_success:    
                        info = {"Report name": sheet_name,'Success':False,'Status' : 'download failed','Info' : run_task_result.msg ,'Sheet_name':sheet_name}
                        download_report_list.append(info)
                    else:
                        info = {"Report name": sheet_name,'Success':True,'Status' : 'download successfully','Info' : '','Sheet_name':sheet_name}
                        download_report_list.append(info)
            
            final_run_task_list = Utility.re_download_report(download_report_list,config_file_path, my_client,2)
            return_value = self.final_result_msg(final_run_task_list,return_value)

            Utility.generate_download_report(final_run_task_list) 
            return_value.width = 800
            logging.info('run_vba_task end')  
        except Exception as ex:
            logging.exception('run_tasks :'+ str(ex))
            return_value.msg ='Download report failed,please check error.log'
            return_value.title = 'Download report failed'
            return_value.is_success = False
        finally:
            deleted_folder = False
            delete_count = 0
            while not deleted_folder:
                try:
                    if delete_count>5:
                        return_value.msg = return_value.msg+'and clear folder failed'
                        break
                    Utility.clear_folder(PublicData.download_report_folder_path)
                    os.rmdir(PublicData.download_report_folder_path)
                    deleted_folder = True
                except Exception as ex:
                    logging.info('rmdir delete folder:'+str(ex))
                    time.sleep(1)
                    delete_count = delete_count+1

            if alert_msg:
                my_client.ExecuteJavascript('hideLoading()')
                my_client.ExecuteJavascript('showMessage("'+return_value.msg+'","'+return_value.title+'","'+str(return_value.width)+'")')
            SeleniumHelper.close_driver(PublicData.web_driver)

            logging.info('generate_download_report start')
            #Utility.generate_download_report(final_run_task_list) 
            logging.info('generate_download_report start')
            Utility.copy_log_info(PublicData.download_report_excel_sheet_name)
            Utility.copy_log_info(PublicData.log_upload_sheet_name)
            Utility.copy_log_info(PublicData.log_approve_sheet_name)
            return return_value

    def final_result_msg(self,final_run_task_list,return_value):
        ''' get html format message frome download result  '''
        try:
            if len(final_run_task_list)> 0:                
                msg = '<table class=\'table\'><tr><th>Report name</th><th>Status</th><th>Note</th><tbody>'
                for item in final_run_task_list:
                    if item['Success'] == False:
                     return_value.is_success = False
                    msg = msg+'<tr><td>'+item['Report name']+'</td><td>'+item['Status']+'</td><td>'+item['Info']+'</td></tr>'
                msg = msg+'</tbody></table>'
                return_value.msg =msg
            else:
                return_value.is_success = False      
            if return_value.is_success:
                return_value.title = 'Download report successfully'                
            else:
                if len(final_run_task_list) == 0:
                    return_value.msg ='Download report failed,please check error.log'
                return_value.title = 'Download report failed'
        except Exception as ex:
            logging.exception('final_result_opreate :'+ str(ex))
        finally:
            return return_value

    def OnSetFocus(self, _):
        if not self.browser:
            return
        if WINDOWS:
            cef.WindowUtils.OnSetFocus(self.browser_panel.GetHandle(),
                                       0, 0, 0)
        self.browser.SetFocus(True)

    def get_task_config_data(self,config_file_path,download_type,report_name):
        try:
            list_config_report_data = []
            if download_type == 1:
                list_config_report_data = Utility.get_report_list(config_file_path,PublicData.config_sheet_name)
            else:
                list_config_report_data_temp = Utility.get_report_list(config_file_path,PublicData.config_sheet_name)
                for row_data in list_config_report_data_temp:
                    if row_data[PublicData.str_excel_config_colunm_report_name] == report_name:
                        list_config_report_data.append(row_data)
                        break
            return list_config_report_data
        except Exception as ex:
            logging.exception('run_tasks :'+ str(ex))
            return list_config_report_data

    def OnSize(self, _):
        if not self.browser:
            return
        if WINDOWS:
            cef.WindowUtils.OnSize(self.browser_panel.GetHandle(),
                                   0, 0, 0)
        elif LINUX:
            (x, y) = (0, 0)
            (width, height) = self.browser_panel.GetSize().Get()
            self.browser.SetBounds(x, y, width, height)
        self.browser.NotifyMoveOrResizeStarted()

    def OnClose(self, event):
        #SeleniumHelper.close_driver(PublicData.web_driver)
        print("[wxpython.py] OnClose called")
        if not self.browser:
            # May already be closing, may be called multiple times on Mac
            return

        if MAC:
            # On Mac things work differently, other steps are required
            self.browser.CloseBrowser()
            self.clear_browser_references()
            self.Destroy()
            global g_count_windows
            g_count_windows -= 1
            if g_count_windows == 0:
                cef.Shutdown()
                wx.GetApp().ExitMainLoop()
                # Call _exit otherwise app exits with code 255 (Issue #162).
                # noinspection PyProtectedMember
                os._exit(0)
        else:
            # Calling browser.CloseBrowser() and/or self.Destroy()
            # in OnClose may cause app crash on some paltforms in
            # some use cases, details in Issue #107.
            self.browser.ParentWindowWillClose()
            event.Skip()
            self.clear_browser_references()

    def clear_browser_references(self):
        # Clear browser references that you keep anywhere in your
        # code. All references must be cleared for CEF to shutdown cleanly.
        self.browser = None

class FocusHandler(object):
    def OnGotFocus(self, browser, **_):
        # Temporary fix for focus issues on Linux (Issue #284).
        if LINUX:
            print("[wxpython.py] FocusHandler.OnGotFocus:"
                  " keyboard focus fix (Issue #284)")
            browser.SetFocus(True)
