from constants import *
from linkSQL import sqlConnect
from util import *
import pandas as pd


""" TIPS
1. 這個版本還不能跨月select (之後導入datetime就可以了，等我一下下OuO)
2. 程式偵測到excel格子裡有數字會加上原本的數字再寫入，如果同一天的資料分散在兩個table不會互相覆蓋
"""


# 要繼續寫入資料的檔案
input_data = "資料統計-路口設備狀態_ping_20220920.xlsx"
# 要執行搜尋的table
db_table_name = "DevicePingStatus"
# sql指令 : OuO_o23是完整24小時的，OuO_x23是扣除23點的
sql_select = OuO_o23
# 西元年
year = 2022
# 月份
mon = 9
# 起始日期
start_day = 20
# 結束日期
end_day = 21
# excel表格的頁籤順序 : 0, 1, 2, ......
xlsx_sheet_index = 2


def find_column(ip, xlsx):
    row = "4"
    for c in range(31):
        cellID = COLUMN[c] + row
        if xlsx[cellID].value == ip:
            return COLUMN[c]


def check_status(status):
    if status == 'T':
        return 1
    if status == 'F':
        return 2


def main_auto():
    try:
        conn, cur = sqlConnect()
        ws, wb = openExcelFile(input_data)
        sheets = wb.sheetnames
        ws = wb[sheets[xlsx_sheet_index]]
        for day in range(start_day, end_day+1):
            print("\nnow select ... {}".format(day), end='')
            set_time_st = "SET @st = '{}-{:02d}-{:02d} 00:00:00.000';".format(year, mon, day)
            cur.execute(set_time_st)
            set_time_ed = "SET @et = '{}-{:02d}-{:02d} 00:00:00.000';".format(year, mon, (day + 1))
            cur.execute(set_time_ed)
            # main
            cur.execute(sql_select.format(tb=db_table_name))
            fetch_data = cur.fetchall()
            df = pd.DataFrame(fetch_data)
            # write data
            if df.empty:
                print("查無資料", end='')
                pass
            else:
                for df_row in range(len(df.index)):
                    row_pingIP = df.iloc[df_row][1]
                    row_status = df.iloc[df_row][2]
                    row_counts = df.iloc[df_row][3]
                    col = find_column(row_pingIP, ws)
                    row = str(4*day + check_status(row_status))
                    cellID = col + row
                    if ws[cellID].value != 0:
                        ws[cellID].value = ws[cellID].value + row_counts
                    else:
                        ws[cellID].value = row_counts
            output_file_name = "資料統計-路口設備狀態_ping_{}{:02d}{:02d}.xlsx".format(year, mon, day)
            wb.save(output_file_name)
            if day < end_day:
                ws, wb = openExcelFile(output_file_name)
                sheets = wb.sheetnames
                ws = wb[sheets[xlsx_sheet_index]]
        conn.close()

        print("\n導入完成OuO")
    except KeyboardInterrupt:
        print("Bye Bye :)")
    except Exception as e:
        print("\n---------------- Error ----------------")
        print("------ Sorry 請截圖下方資訊給Jessie QAQ ------")
        catchError(e)


if __name__ == '__main__':
    main_auto()

