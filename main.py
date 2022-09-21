from constants import *
from linkSQL import sqlConnect
from util import *
import pandas as pd


input_data = "資料統計-路口設備狀態_ping_20220920.xlsx"
db_table_name = "DevicePingStatus"
sql_select = OuO_o23
mon = 9
start_day = 20
end_day = 21
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
        for day in range(start_day, end_day):
            print("\nnow select ... {}".format(day), end='')
            set_time_st = "SET @st = '2022-{:02d}-{:02d} 00:00:00.000';".format(mon, day)
            cur.execute(set_time_st)
            set_time_ed = "SET @et = '2022-{:02d}-{:02d} 00:00:00.000';".format(mon, (day + 1))
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
            wb.save("output_now.xlsx")
            ws, wb = openExcelFile("output_now.xlsx")
            sheets = wb.sheetnames
            ws = wb[sheets[xlsx_sheet_index]]
        wb.save("output.xlsx")
        conn.close()
        print("\n導入完成OuO")
    except KeyboardInterrupt:
        print("Bye Bye :)")
    except Exception as e:
        print("\n---------------- Error ----------------")
        catchError(e)


if __name__ == '__main__':
    main_auto()

