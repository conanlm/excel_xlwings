from string import digits

s = '新中关店2'
remove_digits = str.maketrans('', '', digits)
res = s.translate(remove_digits)
print(res)
