# import xlwings as xw
# pip install pywin32
# pip install PyInstaller
# pyinstaller -F pyexcel.py
import updateexcel as ue
import time

start = time.perf_counter()
# app=xw.App(visible=True,add_book=False)

# book=app.books.open(r'分平台营业数据 (31).xlsx')

# # book = xw.Book('test.xlsx')
# # 此时界面上会弹出Excel窗口，如果test.xlsx文件不存在则会报错，如果test.xlsx已经被打开，直接返回这个文件对象

# # print (book.name,book.fullname)    # 打印文件名和绝对路径
# # print (book.app)    # 可以查看book所在哪个APP
# # print (book.sheets)    # 又是一个类列表结构，存放各种Sheet对象
# book.activate()    # 如果Excel没有获得当前系统的焦点，调用这个方法可以回到Excel中去

# # book.sheets['sheet1'].range('A2').value = 'Foo1'

# # book.sheets['sheet1'].range('A2').value = '123'

# # a=book.sheets['sheet1'].range('A2').expand().value

# # print(a)

# # sht=book.sheets['sheet1']

# # chart = book.sheets['sheet1'].charts.add()
# # chart.set_source_data(sht.range('A1').expand())
# # chart.chart_type = 'line'
# # chart.name

# # book.sheets['王府井店1'].range('S32').value = '61'
# # book.sheets['王府井店1'].range('T32').value = '3580.9'

# arr = []
# # arr.append(book.sheets[0].range('C2').value)
# # arr.append(book.sheets[0].range('D2').value)
# # arr.append(book.sheets[0].range('E2').value)
# # arr.append(book.sheets[0].range('F2').value)
# # arr.append(book.sheets[0].range('G2').value)
# # arr.append(book.sheets[0].range('H2').value)
# arr.append(book.sheets[0].range('I2').value)
# # arr.append(book.sheets[0].range('J2').value)
# # arr.append(book.sheets[0].range('K2').value)
# # arr.append(book.sheets[0].range('L2').value)
# # arr.append(book.sheets[0].range('O2').value)
# # arr.append(book.sheets[0].range('Q2').value)
# # arr.append(book.sheets[0].range('T2').value)
# # arr.append(book.sheets[0].range('U2').value)
# # arr.append(book.sheets[0].range('V2').value)
# # arr.append(book.sheets[0].range('Y2').value)
# # arr.append(book.sheets[0].range('AB2').value)
# # arr.append(book.sheets[0].range('AC2').value)
# # arr.append(book.sheets[0].range('AD2').value)
# # arr.append(book.sheets[0].range('AE2').value)
# # arr.append(book.sheets[0].range('AH2').value)
# # arr.append(book.sheets[0].range('AJ2').value)
# print(arr)
# # book.save()

# # book.app.kill()
# book.close()    # 关闭Excel文档，但只是关闭文件本身，不关闭excel程序。。若要关闭Excel程序则需要调用响应APP实例的kill方法。经过试验，先调用close会导致默认创建的app实例自动消失，从而无法调用kill，从而关不掉Excel
# # 所以最好的办法不是调用这个close而是调用app.kill()。

# # sheet = book.sheets[0]
# # 其他获取sheet对象的方法还有book.sheets['sheet_name']

# book1=app.books.open(r'外送部数据记录表04.25.xls')
# book1.activate()
# # book1.sheets['王府井店1'].range('Y32').value = 123
# book1.sheets['王府井店1'].range('S32').options(numbers=int).value
# # book1.sheets['王府井店1'].range('T32').value = arr[1]
# # book1.sheets['王府井店1'].range('U32').value = arr[2]
# # book1.sheets['王府井店1'].range('X32').value = arr[3]
# # book1.sheets['王府井店1'].range('AA32').value = arr[4]
# # book1.sheets['王府井店1'].range('AB32').value = arr[5]
# # book1.sheets['王府井店1'].range('AE32').value = arr[6]
# # book1.sheets['王府井店1'].range('AG32').value = arr[7]
# # book1.sheets['王府井店1'].range('AH32').value = arr[8]
# # book1.sheets['王府井店1'].range('AJ32').value = arr[8]
# # book1.sheets['王府井店1'].range('AK32').value = arr[9]
# # book1.sheets['王府井店1'].range('AL32').value = arr[10]
# # book1.sheets['王府井店1'].range('AO32').value = arr[11]
# # book1.sheets['王府井店1'].range('AR32').value = arr[12]
# # book1.sheets['王府井店1'].range('AS32').value = arr[13]
# # book1.sheets['王府井店1'].range('AT32').value = arr[14]
# # book1.sheets['王府井店1'].range('AU32').value = arr[15]
# # book1.sheets['王府井店1'].range('AX32').value = arr[16]
# # book1.sheets['王府井店1'].range('AZ32').value = arr[17]

# book1.save()
# book1.close()
# print ()
a = []
for line in open("name.txt"):
    a.append(line.strip())

# a.append(open("name.txt",'r', encoding='UTF-8').read() )

# arr.add
print(a)

# 分平台营业数据 (31).xlsx
# 评论率 (30).xlsx
# 外送部数据记录表04.25.xls
# 32
# 93

ue.update(a)

end = time.perf_counter()
print(end - start)
