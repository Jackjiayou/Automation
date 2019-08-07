import PublicData
import logging
import time

def add_log_by_logging(msg):
    """add log."""
    try:       
        logging.info(msg)
    except Exception as ex:
        with open(PublicData.app_log_file_path,'a') as f:
            f.write(time.strftime('%Y.%m.%d-%H%M%S',time.localtime(time.time()))+':'+str(ex)+'\n')   
            f.close()

def add_log(msg):
    """add log."""
    try:       
        with open(PublicData.app_log_file_path,'a') as f:
            f.write(time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))+' INFO: '+str(msg)+'\n')               
    except Exception as ex:
        logging.exception('add_log :'+str(ex))
    finally:
        f.close()

def config_log_info(file_path):
    logging.basicConfig(level=logging.INFO,#控制台打印的日志级别
                    filename=file_path,
                    filemode='a',##模式，有w和a，w就是写模式，每次都会重新写日志，覆盖之前的日志
                    #a是追加模式，默认如果不写的话，就是追加模式
                    format=
                    '%(asctime)s - %(pathname)s[line:%(lineno)d] - %(levelname)s: %(message)s'
                    #日志格式
                    )
