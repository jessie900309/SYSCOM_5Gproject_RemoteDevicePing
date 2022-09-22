import datetime
import sys
import traceback
import openpyxl


def catchError(e):
    cl, exc, tb = sys.exc_info()
    lastCallStack = traceback.extract_tb(tb)
    fileName = lastCallStack[0]
    print(fileName)
    print(e)
    sys.exit(1)


def openExcelFile(op):
    wb = openpyxl.load_workbook(op, data_only=False)
    wb.active = 0
    ws = wb.active
    return ws, wb


def getDateList(start, end):
    st = datetime.date(start[0], start[1], start[2])
    et = datetime.date(end[0], end[1], end[2])
    if st > et:
        raise ValueError
    else:
        numdays = et - st
        date_list = []
        for x in range(numdays.days + 2):
            date_list.append((st + datetime.timedelta(days=x)).strftime('%Y-%m-%d'))
        return date_list

