# 载入库
import xlrd
import xlwt
from xlwt import Workbook
import numpy as np

# 载入原始Excel表
namein = input('请输入原始数据文件名称\n')
loc = (namein + '.xlsx')

# 打开表格及工作簿
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

# 行数
row = sheet.nrows

# 载入膜号列数
col = int(input('请输入膜号所在列数，请用数字表示\n')) - 1
col_of_index = [col]

# 录入T带位置、C带位置等
# 用新的dist录入C带T带的值和位置
b = []
value = []
index = -1
j = col_of_index
for i in range(row):
    cur = sheet.cell_value(i, j)
    if type(cur) == float:
        for k in range(1, 5):
            rValue = sheet.cell_value(i, j + k)
            if type(rValue) == float or type(rValue) == int:
                if index < 0 or b[index][0] != int(cur):
                    b.append((int(cur), [rValue]))
                    index = index + 1
                else:
                    b[index][1].append(rValue)

wb = Workbook()
sheet1 = wb.add_sheet('不合格膜号')
sheet2 = wb.add_sheet('合格膜号')
downT = float(input('T带位置的下限是\n'))
upT = float(input('T带位置的上限是\n'))
downC = float(input('C带位置的下限是\n'))
upC = float(input('C带位置的上限是\n'))
absT = float(input('T带值的下限是\n'))
absC = float(input('C带值的下限是\n'))
el = int(input('初筛还是复筛？\n初筛请输入1，复筛请输入2\n'))  # 区分初筛、复筛
er = 0
n = 1
m = 1
for i in range(index + 1):
    if b[i][1][1] >= 0.15:  # 筛选CV过大
        er += 2
        if er >= el:
            sheet1.write(n, 0, b[i][0])
            i += 1
            n += 1
        for j in range(4, 8):  # 筛选T位置偏移
            if downT <= b[i][1][j] <= upT:
                er += 1
                if er >= el:
                    sheet1.write(n, 0, b[i][0])
                    i += 1
                    n += 1
        for j in range(8, 12):  # 筛选C位置偏移
            if downC <= b[i][1][j] <= upC:
                er += 1
                if er >= el:
                    sheet1.write(n, 0, b[i][0])
                    i += 1
                    n += 1
        for j in range(12, 16):  # 筛选T绝对值低
            if b[i][1][j] <= absT:
                er += 1
                if er >= el:
                    sheet1.write(n, 0, b[i][0])
                    i += 1
                    n += 1
        for j in range(16, 20):  # 筛选C绝对值低
            if b[i][1][j] <= absT:
                er += 1
                if er >= el:
                    sheet1.write(n, 0, b[i][0])
                    i += 1
                    n += 1

# 导出工作簿
name = input('请输入文件名称\n叫 不合格膜号 可以么？\n请输入 \"可以\" 或者 文件名称\n')
if name == '可以':
    wb.save('不合格膜号.xls')
else:
    wb.save(name + '.xls')

# 结束提示
print('请查看文件')
