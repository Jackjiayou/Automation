#!/usr/bin/env python    # -*- coding: utf-8 -*
import Utility
from selenium import webdriver
import os
import time
from return_value_model import return_data
import logging
import log_helper

def create_driver(web_Driver,currentUserDir,myRootPath,browser_type,download_file_path,hide_window = False,read_user_root_folser = True):     
    ''' create browser '''
    try:
        #browser_type = 'Ie'
        if not web_Driver == None and web_Driver != False:
            close_driver(web_Driver)
        if browser_type == 'Ie':
            web_Driver = webdriver.Ie()   
        elif browser_type == 'chrome':
            driverOptions = webdriver.ChromeOptions()
            prefs = {"download.default_directory":download_file_path}
            driverOptions.add_experimental_option("prefs", prefs)

            if read_user_root_folser:
                driverOptions.add_argument(r"--user-data-dir="+currentUserDir)
            #if hide_window:
            #    driverOptions.add_argument('--headless')
            #    driverOptions.add_argument('--disable-gpu')   

            chromedriver = myRootPath+'\chromedriver.exe'
            os.environ["webdriver.chrome.driver"] = chromedriver
            web_Driver = webdriver.Chrome(chromedriver,0,driverOptions)    
        return web_Driver          
    except Exception as ex:
        logging.exception("create_browser :"+str(ex))
        close_driver(web_Driver)
        return None

def quit_driver(web_Driver):
    try:
        web_Driver.quit()
    except Exception as ex:
        logging.exception("quit_driver"+str(ex))

def close_driver(web_Driver):
    try:
        if web_Driver != None:
            web_Driver.close()
            web_Driver.quit()
    except Exception as ex:
        logging.exception("disable_browser"+str(ex))
        quit_driver(web_Driver)
    finally:
        web_Driver = None      
