import xlwings as xw

from string import digits
import time
start = time.perf_counter()

# app=xw.App(visible=True,add_book=False)

# book=app.books.open(r'test.xlsx')

# arr=book.sheets[0].range('C2:v2').value

# print(arr)


# book1=xw.Book('外送部数据记录表04.25.xls')
# book=xw.Book('分平台营业数据 (31).xlsx')
# # 获取表名
# print (book1.sheets[2].name)
# # 表总量
# print (len(book1.sheets))
# # 表总列数
# print (book.sheets[0].used_range.last_cell.row)
# # B2一列的值
# print (book.sheets[0].range('B2').expand('down').value)

# # 去掉字符串中的数字
# # s = '新中关店2'
# # remove_digits = str.maketrans('', '', digits)
# # res = s.translate(remove_digits)
# # print(res)
# i=1
# arr=[]
# while(i<(len(book1.sheets)-1)) :
#     remove_digits = str.maketrans('', '', digits)
#     res = book1.sheets[i].name.translate(remove_digits)

#     arr.append(res)
#     i=i+1

# print(arr)


# book.close()


def update():
    # app=xw.App(visible=True,add_book=False)

    # book=app.books.open(r'分平台营业数据 (31).xlsx')
    # print(a[2])

    #  time.sleep(3)

    #  book=app.books.open(a[0])
    #  book1=app.books.open(a[2])
    #  book2=app.books.open(a[1])

    try:
        book = xw.Book('分平台营业数据 (34).xlsx')
        book1 = xw.Book('外送部数据记录表04.28.xls')
        book2 = xw.Book('评论率 (33).xlsx')
        book3 = xw.Book('伏牛堂-日推广消费（2.0）.xlsx')
    except FileNotFoundError:
        print("二孩子，文件名写对了吗")
    else:

        print(book.app)    # 可以查看book所在哪个APP
        print(book.sheets)    # 又是一个类列表结构，存放各种Sheet对象

        #  time.sleep(3)

        # book1=app.books.open(r'外送部数据记录表04.25.xls')
        # book2=app.books.open(r'评论率 (30).xlsx')
        # book=app.books.open(r'test.xlsx')

        # book = xw.Book('test.xlsx')
        # 此时界面上会弹出Excel窗口，如果test.xlsx文件不存在则会报错，如果test.xlsx已经被打开，直接返回这个文件对象

        # print (book.name,book.fullname)    # 打印文件名和绝对路径
        # print (book.app)    # 可以查看book所在哪个APP
        # print (book.sheets)    # 又是一个类列表结构，存放各种Sheet对象
        # book.activate()    # 如果Excel没有获得当前系统的焦点，调用这个方法可以回到Excel中去

        # book.sheets['sheet1'].range('A2').value = 'Foo1'

        # book.sheets['sheet1'].range('A2').value = '123'

        # a=book.sheets['sheet1'].range('A2').expand().value

        # print(a)

        # sht=book.sheets['sheet1']

        # chart = book.sheets['sheet1'].charts.add()
        # chart.set_source_data(sht.range('A1').expand())
        # chart.chart_type = 'line'
        # chart.name

        # book.sheets['王府井店1'].range('S32').value = '61'
        # book.sheets['王府井店1'].range('T32').value = '3580.9'

        i = 2
        arr5 = book3.sheets[0].range('D3:D32').value
        arr1 = []
        arr2 = []
        arr3 = []
        arr4 = []
        while(i < 32):
            # arr1 = book.sheets[0].range('C'+str(i)+':E'+str(i)).value
            # arr2 = book.sheets[0].range('K'+str(i)+':V'+str(i)).value
            # arr3 = book.sheets[0].range('AB'+str(i)+':AK'+str(i)).value
            # arr4 = book2.sheets[0].range('C'+str(i)+':E'+str(i)).value
            arr1.append(book.sheets[0].range('C'+str(i)+':E'+str(i)).value)
            arr2.append(book.sheets[0].range('K'+str(i)+':V'+str(i)).value)
            arr3.append(book.sheets[0].range('AB'+str(i)+':AK'+str(i)).value)
            arr4.append(book2.sheets[0].range('C'+str(i)+':E'+str(i)).value)

            # print(arr1)
            # print(arr2)
            # book.sheets[0].range('J8').value=88.8
            # book.sheets[0].range('J8').api.Font.Bold = True
            # book.sheets[0].range('J8').api.NumberFormatLocal= "0.00_);[红色](0.00)"
            # book.save()

            # # book.app.kill()
            # 关闭Excel文档，但只是关闭文件本身，不关闭excel程序。。若要关闭Excel程序则需要调用响应APP实例的kill方法。经过试验，先调用close会导致默认创建的app实例自动消失，从而无法调用kill，从而关不掉Excel
            # # 所以最好的办法不是调用这个close而是调用app.kill()。

            # # sheet = book.sheets[0]
            # # 其他获取sheet对象的方法还有book.sheets['sheet_name']

            #   book1.activate()

            print(i)
            i = i+1
        t = 2
        row = 32
        row1 = 93
        i = 2
        while(i < 32):

            book1.sheets[i-1].range('S'+str(row)).value = arr1[i-2]
            book1.sheets[i-1].range('AA'+str(row)).value = arr2[i-2]
            book1.sheets[i-1].range('AR'+str(row)).value = arr3[i-2]

            book1.sheets[i-1].range('F'+str(row1)).value = arr4[0]
            book1.sheets[i-1].range('H'+str(row1)).value = arr4[1]
            book1.sheets[i-1].range('J'+str(row1)).value = arr4[2]

            book1.sheets[i-1].range('Q'+str(row)).value = arr5[i-2]

            book1.sheets[i-1].range('AJ'+str(row)
                                    ).api.NumberFormatLocal = "G/通用格式"
            # book1.sheets['王府井店1'].range('AK32').api.NumberFormatLocal= "0.00_);[红色](0.00)"
            book1.sheets[i-1].range('AK'+str(row)
                                    ).api.NumberFormatLocal = "G/通用格式"
            # book1.sheets['王府井店1'].range('AL32').api.Font.Bold
            book1.sheets[i-1].range('AL'+str(row)
                                    ).api.NumberFormatLocal = "G/通用格式"
            print(i)
            i = i+1

        # print (book1.sheets)

        book1.save()
        book.app.kill()
    #  book1.app.kill()
    #  book2.app.kill()


update()
end = time.perf_counter()
print(end-start)
