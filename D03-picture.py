# -*- coding: utf-8 -*-
# @Time    : 2020/5/18 12:14
# @Author  : 结尾！！
# @FileName: D03-picture.py
# @Software: PyCharm



import pymongo
import requests

def open_requests(asin,img):
    img_url = f'https://images-na.ssl-images-amazon.com/images/I/{img}'
    res=requests.get(img_url)
    with open(f"./downloads_picture/{asin}.jpg",'wb') as fn:
        fn.write(res.content)



conn=pymongo.MongoClient(host='127.0.0.1',port=27017)
db=conn.operate_2020
db_table=db['keepa_data']  #连接爬的变体的数据库




import xlwings as xw

app = xw.App(visible=False, add_book=False)
wb = app.books.open('./文件/美国腰包 fanny pack20200518.xlsx')
sht_name = wb.sheets[1].name
print(wb.sheets)
sht=wb.sheets[sht_name]

ASIN_LIST=sht.range('C2').expand('down').value
wb.close()
app.quit()
# print(ASIN_LIST)

#查看已经保存的图片
import os
file_list = os.listdir('./downloads_picture')
file_list_asin=[file[:10]   for file in file_list  if file.endswith('.jpg')]
print(file_list_asin)

for index,ASIN in enumerate(ASIN_LIST):

    info=db_table.find_one({"asin":ASIN},{'_id':0})
    if info:
        img=info['imagesCSV'].split(',')
        print(img)
        if ASIN not in file_list_asin:
            open_requests(ASIN,img[0])
    else:
        print(ASIN,'没有')

'''1416958819.0 没有'''



