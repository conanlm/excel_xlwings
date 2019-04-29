import xlwings as xw

from string import digits

# app=xw.App(visible=True,add_book=False)

# book=app.books.open(r'test.xlsx')

# arr=book.sheets[0].range('C2:v2').value

# print(arr)

book1=xw.Book('外送部数据记录表04.25.xls')
book=xw.Book('分平台营业数据 (31).xlsx')
# 获取表名
print (book1.sheets[2].name)     
# 表总量           
print (len(book1.sheets))
# 表总列数
print (book.sheets[0].used_range.last_cell.row)
# B2一列的值
print (book.sheets[0].range('B2').expand('down').value)

# 去掉字符串中的数字
# s = '新中关店2'
# remove_digits = str.maketrans('', '', digits)
# res = s.translate(remove_digits)
# print(res)
i=1
arr=[]
while(i<(len(book1.sheets)-1)) :
    remove_digits = str.maketrans('', '', digits)
    res = book1.sheets[i].name.translate(remove_digits)

    arr.append(res)
    i=i+1

print(arr)


book.close() 

