#/Users/Dean/Desktop/t.xls
import xlrd
import sys



path = sys.argv[1]
# workbook = xlrd.open_workbook('/Users/Dean/Desktop/d.xlsx')
workbook = xlrd.open_workbook(path)

sheet2 = workbook.sheets()[0]

start = 1
end = 0
start_date = 0
nrows = sheet2.nrows

first_date = xlrd.xldate_as_tuple(sheet2.cell(start , 1).value, 0)[0:2]

l = []
l2 = []
all_cnt = 0
for i in range(2, nrows):
    if xlrd.xldate_as_tuple(sheet2.cell(i , 1).value, 0)[0:2] != first_date[0:2]:
        end = i - 1
        l.append((start, end))
        start = i
        first_date = xlrd.xldate_as_tuple(sheet2.cell(start , 1).value, 0)[0:2]
    if i == nrows - 1:
        end = nrows - 1
        l.append((start, end))


for i in l:
    cnt = 0
    for j in range(i[0], i[1] + 1):
        v7 = float(sheet2.cell(j, 7).value)
        v8 = float(sheet2.cell(j, 8).value)
        cnt += v7  - 0.5
        if v7 == 0.0:
            cnt -= 8.5

    all_cnt += round(cnt)
    l2.append((xlrd.xldate_as_tuple(sheet2.cell(i[0] , 1).value, 0)[0], xlrd.xldate_as_tuple(sheet2.cell(i[0] , 1).value, 0)[1], round(cnt)))


print(l2)





