import pandas as pd
from excel_cell_func import *
from linkSQL import sqlConnect
from util import *


""" TIPS
1. 沒辦法自動換db的table、excel表格的頁籤
2. 程式偵測到excel格子裡有數字會加上原本的數字再寫入，如果同一天的資料分散在兩個table不會互相覆蓋
3. 如果執行到一半程式死掉，ouo資料夾裡面的output_now是程式死掉的前一天的資料
(比如說select 8/24 的時候網路斷掉導致程式死掉，output_now就會是8/23之前的資料)
"""

"""
設定參數 : 
"""

# 要繼續寫入資料的檔案(跟main.py在不同目錄的話要寫完整路徑喔OuO)
input_data = "資料統計-路口設備狀態_ping_2022-09-20.xlsx"
# 要執行搜尋的table
db_table_name = "DevicePingStatus_20220920"
# sql指令 : OuO_o23是完整24小時的，OuO_x23是扣除23點的
sql_select = OuO_o23
# 起迄日期 : 西元年, 月, 日
start_day = [2022, 9, 18]
end_day = [2022, 9, 20]
# excel表格的頁籤順序 : 0, 1, 2, ......
xlsx_sheet_index = 2

"""
設定完參數直接執行這個main.py就可以了OuO
"""


def main():
    try:
        select_date_list = getDateList(start_day, end_day)
        print(select_date_list)
        conn, cur = sqlConnect()
        ws, wb = openExcelFile(input_data)
        sheets = wb.sheetnames
        ws = wb[sheets[xlsx_sheet_index]]
        for day in range(len(select_date_list)-1):
            print("\nnow select ... {}".format(select_date_list[day]), end='')
            set_time_st = "SET @st = '{} 00:00:00.000';".format(select_date_list[day])
            cur.execute(set_time_st)
            set_time_ed = "SET @et = '{} 00:00:00.000';".format(select_date_list[day+1])
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
                    row = str(4*int(select_date_list[day][-2:]) + check_status(row_status))
                    cellID = col + row
                    if ws[cellID].value != 0:
                        ws[cellID].value = ws[cellID].value + row_counts
                    else:
                        ws[cellID].value = row_counts
            wb.save("ouo/output_now.xlsx")
            ws, wb = openExcelFile("ouo/output_now.xlsx")
            sheets = wb.sheetnames
            ws = wb[sheets[xlsx_sheet_index]]
        if sql_select == OuO_x23:
            filetype = "扣除23時_"
        else:
            filetype = ""
        wb.save("資料統計-路口設備狀態_ping_" + filetype + select_date_list[-2] + ".xlsx")
        conn.close()
        print("\n導入完成OuO")
    except KeyboardInterrupt:
        print("\nBye Bye :)")
    except ValueError:
        print("\n你輸入的日期不合理喔OHO")
    except Exception as e:
        print("\n\n---------------- Error ----------------")
        print("---- Sorry 請截圖下方資訊給Jessie QAQ ----")
        catchError(e)


if __name__ == '__main__':
    main()

