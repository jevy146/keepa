# -*- coding: utf-8 -*-
# @Time    : 2020/5/18 11:51
# @Author  : 结尾！！
# @FileName: D02_write_change_asin.py
# @Software: PyCharm


import pymongo


conn=pymongo.MongoClient(host='127.0.0.1',port=27017)
db=conn.operate_2020
db_table=db['keepa_data']  #连接爬的变体的数据库



import xlwings as xw

app = xw.App(visible=False, add_book=False)
wb = app.books.open('./文件/美国腰包 fanny pack20200518.xlsx')
sht_name = wb.sheets[1].name
sht=wb.sheets[sht_name]

ASIN_LIST=sht.range('C2').expand('down').value

# print(ASIN_LIST)


for index,ASIN in enumerate(ASIN_LIST):


    info=db_table.find_one({"asin":ASIN},{'_id':0})
    if  info:
        print(info['variationCSV'])
        sht.range(f'D{index + 2}').value = info['variationCSV']
    else:
        sht.range(f'D{index + 2}').value = "null"

wb.save()
wb.close()
app.quit()