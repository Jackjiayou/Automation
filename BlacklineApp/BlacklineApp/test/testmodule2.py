import openpyxl
import log_helper
from openpyxl import load_workbook
import os
import Utility

def is_not_open(filename):
    try:
        wb = openpyxl.load_workbook(filename)
        wb.save(filename)
        return False
        #if not os.access(filename, os.F_OK):
        #    return True # file doesn't exist
        #h = _sopen(filename, 0, _SH_DENYRW, 0)
        #if h == 3:
        #    _close(h)
        #    logging.info('is_open open:' +filename)
        #    return True # file is not opened by anyone else
        #return False # file is already open
    except Exception as ex:
        logging.exception('is_open errr:' +str(ex))
        return True

path = r'C:\Blackline Automation App\Report\July\EALA\ImportGroup\AngolaAccenture_236XXX_Jul19.xlsx'

if os.path.isfile(path):

    print(123)