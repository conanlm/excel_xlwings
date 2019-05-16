# pip install xlwings
import xlwings as xw
# import time


def update(a):
    # app=xw.App(visible=True,add_book=False)

    # book=app.books.open(r'分平台营业数据 (31).xlsx')
    # print(a[2])

    #  time.sleep(3)

    #  book=app.books.open(a[0])
    #  book1=app.books.open(a[2])
    #  book2=app.books.open(a[1])
    # excel 默认设置：程序可见，只打开不新建工作薄，屏幕更新关闭
    app = xw.App(visible=False, add_book=False)
    # app.display_alerts=False
    # app.screen_updating=False
    try:
        book = app.books.open(str(a[0]))
        book1 = app.books.open(str(a[2]))
        book2 = xw.Book(str(a[1]))
        book3 = xw.Book(a[5])
        book4 = xw.Book(a[6])
        book5 = xw.Book(a[7])
    except FileNotFoundError:
        print("二孩子，文件名写对了吗")
        book1.app.quit()
        input()
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
        arr7 = book5.sheets[0].range('B3:B32').value
        arr8 = book5.sheets[0].range('C3:C32').value
        while(i < 32):

            arr1 = book.sheets[0].range('C'+str(i)+':E'+str(i)).value
            # arr2 = book.sheets[0].range('K'+str(i)+':V'+str(i)).value
            arr2 = book.sheets[0].range('K'+str(i)+':S'+str(i)).value
            arr6 = book.sheets[0].range('T'+str(i)+':V'+str(i)).value
            arr3 = book.sheets[0].range('AB'+str(i)+':AK'+str(i)).value
            arr4 = book2.sheets[0].range('C'+str(i)+':E'+str(i)).value
            # print(arr1)
            # print(arr2)
            print(arr1)
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
            row = a[3]
            row1 = a[4]
            row2=int(row)+1
            book1.sheets[i-1].range('Y'+str(row)).value = arr1
            book1.sheets[i-1].range('AG'+str(row)).value = arr2
            book1.sheets[i-1].range('AY'+str(row)).value = arr6
            book1.sheets[i-1].range('BG'+str(row)).value = arr3

            book4.sheets[i].range('F'+str(row1)).value = arr4[0]
            book4.sheets[i].range('H'+str(row1)).value = arr4[1]
            book4.sheets[i].range('J'+str(row1)).value = arr4[2]

            book1.sheets[i-1].range('W'+str(row)).value = arr5[i-2]
            book1.sheets[i-1].range('AP'+str(row2)).value = arr7[i-2]
            book1.sheets[i-1].range('BR'+str(row2)).value = arr8[i-2]

            # 推广格式设置   自定义格式 "[=0]"""";G/通用格式"
            # book1.sheets[i-1].range('W'+str(row)).api.NumberFormatLocal= "0.00_ "

            # book1.sheets[i-1].range('AP'+str(row)
            #                         ).api.NumberFormatLocal = "G/通用格式"
            # # book1.sheets['王府井店1'].range('AK32').api.NumberFormatLocal= "0.00_);[红色](0.00)"
            # book1.sheets[i-1].range('AQ'+str(row)
            #                         ).api.NumberFormatLocal = "G/通用格式"
            # # book1.sheets['王府井店1'].range('AL32').api.Font.Bold
            # book1.sheets[i-1].range('AR'+str(row)
            #                         ).api.NumberFormatLocal = "G/通用格式"
            print(i)
            i = i+1

        # print (book1.sheets)

        book1.save()
        book4.save()
        # wps需要逐个关闭，微软office只需关闭一次，就会全部关闭
        book.app.quit()

        # book1.app.quit()
        # book2.app.quit()
        # book3.app.quit()

        # xw.apps[xw.apps.keys()].quit()
        # print(xw.apps.keys())
        # book.close()
        # book1.close()
        # book2.close()
        # book3.close()

    #  book1.app.kill()
    #  book2.app.kill()
