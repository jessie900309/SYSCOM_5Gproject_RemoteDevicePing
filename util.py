import sys
import traceback
import openpyxl


def catchError(e):
    error_class = e.__class__.__name__
    detail = e.args[0]
    cl, exc, tb = sys.exc_info()
    lastCallStack = traceback.extract_tb(tb)
    fileName = lastCallStack[0]
    # lineNum = lastCallStack[1]
    errMsg = "File \"{}\", \n [{}] {}\n".format(fileName, error_class, detail)
    print(errMsg)
    sys.exit(1)


def openExcelFile(op):
    wb = openpyxl.load_workbook(op, data_only=False)
    wb.active = 0
    ws = wb.active
    return ws, wb


