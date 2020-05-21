# -*- coding: utf-8 -*-
# @Time    : 2020/1/19 10:17
# @Author  : 结尾！！
# @FileName: 4将图片插入到excel中.py
# @Software: PyCharm



from PIL import Image
import os
import xlwings as xw
app=xw.App(visible=True,add_book=False)

wb = xw.Book('./文件/美国腰包 fanny pack20200518.xlsx')
#居中 插入
import os
sht = wb.sheets['美国腰包 fanny pack20200518']
asin_list=sht.range("AA2").expand('down').value
print(len(asin_list))

def write_pic(cell,asin):
    path=f'./downloads_picture/{asin}.jpg'
    print(path)
    fileName = os.path.join(os.getcwd(), path)
    img = Image.open(path).convert("RGB")
    print(img.size)
    w, h = img.size
    x_s = 70  # 设置宽 excel中，我设置了200x200的格式
    y_s = h * x_s / w  #  等比例设置高
    sht.pictures.add(fileName, left=sht.range(cell).left, top=sht.range(cell).top, width=x_s, height=y_s)


if __name__ == '__main__':

    for index,value in enumerate(asin_list):
        cell="A"+str(index+2)
        try:
            write_pic(cell,value)
        except:
            print("没有找到这个ASIN的图片",value)


    wb.save()

