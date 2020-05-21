# -*- coding: utf-8 -*-
# @Time    : 2020/5/18 9:27
# @Author  : 结尾！！
# @FileName: D01data_from_keepa.py
# @Software: PyCharm


import keepa
import matplotlib.pyplot as plt
from datetime import datetime
from keepa.interface import keepa_minutes_to_time
import numpy as np
plt.rcParams['font.sans-serif']=['SimHei'] #用来正常显示中文标签
plt.rcParams['axes.unicode_minus']=False #用来正常显示负号


def keepa_data(ASIN,COUNTRY):
    #https://keepa.com/#!api  keepa数据的api
    accesskey = '8ga5ilhbo9qgknjgsh0ffnbrbbrh6f7742c7ev0f3ge3pp95t5dgvoaepnf66pc7' # enter real access key here
    api = keepa.Keepa(accesskey)
    products = api.query(ASIN,domain=COUNTRY)  #offers=[20-100]
    product = products[0]

    # keepa.plot_product(product)
    ###  直接将所有的数据写入到MongoDB中。
    import pymongo
    # 创建MongoDB数据库连接
    mongocli = pymongo.MongoClient(host = "127.0.0.1", port = 27017)

    # 创建mongodb数据库名称
    dbname = mongocli["operate_2020"]
    # keepa_data 按照获取数据的nowtime进行保存
    sheetname = dbname["keepa_data"] #类目


    '''排名变化'''
    SALES_time=product["data"]["SALES_time"].astype(str)
    SALES=product["data"]["SALES"]
    SALES[SALES<0]=0
    SALES=SALES.astype(str)

    '''价格变化'''
    # AMAZON_time=product['data']['AMAZON_time'].astype(str)  # 价格的日期
    # AMAZON=product["data"]['AMAZON'].astype(str),  # 价格的变化
    #

    now_time = datetime.now() #程序运行的日期
    listed_date =keepa_minutes_to_time(product['listedSince'])  #上架时间
    data_insert={ 'SALES_time':list(SALES_time), #时间
                 'SALES':list(SALES) ,#排名,
                 'nowtime':str(now_time.date()), #写入数据的时间
                 'categoryTree':product["categoryTree"],#类目树
                 'salesRanks':product["salesRanks"],#类目排名
                  'variationCSV':product["variationCSV"], #变体
                  'variations':product["variations"],  #变体属性，变体名称和变体的颜色
                 'asin':product["asin"], #搜索的ASIN
                 'parentAsin':product["parentAsin"] ,#父ASIN
                 'imagesCSV':product["imagesCSV"] ,#图片
                 'title':product["title"] ,#标题
                 'brand':product["brand"] ,#品牌名称
                 'frequentlyBoughtTogether':product["frequentlyBoughtTogether"] ,#经常组合购买
                 'productGroup':product["productGroup"] ,#产品 分组
                 'partNumber':product["partNumber"] ,#零件号  2L
                 'listed_date':listed_date , #上架日期
                 'binding':product["binding"] ,#绑定类目
                 'AMAZON_time':list(product['data']['AMAZON_time'].astype(str)) ,#价格的日期
                 'AMAZON':list(product["data"]['AMAZON'].astype(str) ),#价格的变化
                }
    print(data_insert)
    sheetname.insert_one(data_insert)



import xlwings as xw

app = xw.App(visible=False, add_book=False)
wb = app.books.open('./文件/美国腰包 fanny pack20200518.xlsx')
sht_name = wb.sheets[0].name
sht=wb.sheets[sht_name]

print(sht.range('Z1').value)
ASIN_LIST=sht.range('Z2').expand('down').value
wb.close()
app.quit()
print(len(ASIN_LIST))

import time
# Valid values: [ 1: com | 2: co.uk | 3: de | 4: fr | 5:
#                 co.jp | 6: ca | 7: cn | 8: it | 9: es | 10: in | 11: com.mx ]
DCODES = ['RESERVED', 'US', 'GB', 'DE', 'FR', 'JP', 'CA', 'CN', 'IT', 'ES',
          'IN', 'MX']
if __name__ == '__main__':
    count=0
    for each in ASIN_LIST:
        print(each)
        try:
            keepa_data(each,'US')
            time.sleep(1.2)
        except Exception as e :
            print('错误',e)
        count+=1
        print('*'*10,count)