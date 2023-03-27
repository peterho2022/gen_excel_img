# -*- encoding: utf-8 -*-
'''
@File    :   main.py
@Time    :   2023/03/27 21:45:16
@Author  :   Peter
此程式目的是自動產生Excel圖表，因為我覺得每天做晨報很煩==呵呵
'''

# for csv lib
import csv
# for excel lib
import openpyxl
# 時間函式
# 觀眾說要+pytz代表台北時區(晚點查)
import time


# print(time.strftime("%Y-%m-%d", time.localtime()))
csv_path = "./test.csv"



def read_pasing_csv(csv_path):
    # 計算客戶數量
    customer_cnt = 0
    with open(csv_path, newline='') as csvfile:
        rows = csv.reader(csvfile)

        for row in rows:        
            date = row[0]
            # 因為CSV前面有空格將空格去掉
            customer = row[1].replace(" ", "")
            # 生產日期是否為今天
            if date == time.strftime("%Y-%m-%d", time.localtime()):
                # 客戶是否為 Peter
                if customer == "Peter":
                    customer_cnt = customer_cnt + 1
    return customer_cnt

if __name__ == "__main__":
    # print("總共客戶數量: ", read_pasing_csv(csv_path))
    
    today = time.strftime("%Y-%m-%d", time.localtime())
    customer_cnt = read_pasing_csv(csv_path)

    # 開啟
    workbook = openpyxl.Workbook()

    sheet = workbook.worksheets[0]

    sheet["A2"] = "客戶數"
    sheet['B1'] = today
    sheet['B2'] = customer_cnt

    workbook.save('./test.xlsx')