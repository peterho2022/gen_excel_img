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

import pandas as pd

def read_pasing_csv(csv_path):
    """
    Args:
        csv_path (str) : CSV 路徑
    Returns:
        customer_cnt (int) : 特定用戶的今日數量
    """
    # 初始化客戶數量
    customer_cnt = 0
    df = pd.read_csv(csv_path)
    df = df.reset_index()
    for index, row in df.iterrows():
        date, customer = row["date"], row["customer"]
        # 生產日期是否為今天
        # if date == time.strftime("%Y-%m-%d", time.localtime()):
        # 先固定日期
        if date == "2023-03-27":                
            # 客戶是否為 Peter
            if customer == "Peter":
                customer_cnt = customer_cnt + 1
    return customer_cnt

def main():
    # csv 路徑
    csv_path = "./test.csv"
    # print("總共客戶數量: ", read_pasing_csv(csv_path))
    today = time.strftime("%Y-%m-%d", time.localtime())
    customer_cnt = read_pasing_csv(csv_path=csv_path)

    # 實例化物件
    workbook = openpyxl.Workbook()

    sheet = workbook.worksheets[0]

    sheet["A2"] = "客戶數"
    sheet['B1'] = today
    sheet['B2'] = customer_cnt
    # 輸出路徑
    excel_output_path = "./test.xlsx"
    workbook.save(excel_output_path)

if __name__ == "__main__":
    main()
    
