# 载入库
import xlrd
import xlwt
from xlwt import Workbook
import numpy as np

# 载入原始Excel表
loc = 'test.xlsx'

# 打开表格及工作簿
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

# 行数
row = sheet.nrows

# 载入膜号列数
col_of_index = [2]

# 数据处理、转置、编组
d = []
index = -1
for j in col_of_index:
    for i in range(row):
        cur = sheet.cell_value(i, j)
        if type(cur) == float:
            rValue = sheet.cell_value(i, j + 5)
            if type(rValue) == float or type(rValue) == int:
                if index < 0 or d[index][0] != int(cur):
                    d.append((int(cur), [rValue]))
                    index += 1
                else:
                    d[index][1].append(rValue)

# 出错膜号
error = 0
s = set()
wrong_pos = set()
for i in range(index + 1):
    if d[i][0] in s:
        wrong_pos.add(d[i][0])
    else:
        s.add(d[i][0])

# 将处理好的数据保存
# 创建工作簿
wb = Workbook()
sheet1 = wb.add_sheet('数据处理')
sheet2 = wb.add_sheet('出错数据')

# 标色出错膜号
styleWrongPos = xlwt.easyxf('pattern: pattern solid, fore_colour ocean_blue; font: bold on;')
styleMoreOrLess = xlwt.easyxf('pattern: pattern solid, fore_colour red; font: bold on;')

# 写入抬头
sheet1.write(0, 0, '膜号')
sheet1.write(0, 1, 'RI值')
sheet1.write(0, 2, 'RI值')
sheet1.write(0, 3, 'RI值')
sheet1.write(0, 4, 'RI值')
sheet1.write(0, 5, 'AVG')
sheet1.write(0, 6, 'CV')
sheet2.write(0, 0, '重复膜号')
sheet2.write(0, 2, '缺失膜号')
sheet2.write(0, 4, '不合格膜号')

# 先进行错误膜号的修改
useless_pos = set()
for i in range(len(d)):
    if i > 0:  # 输出错号,修改录入出错的情况再进行录入
        a = len(d[i - 1][1])
        b = len(d[i][1])
        if a + b == 4:  # 修改录入错误的膜号
            for k in range(a - 1, -1, -1):
                d[i][1].append(d[i - 1][1][k])
                del d[i - 1][1][k]
            d[i - 1][1].append(-1)
            useless_pos.add(d[i - 1][0])
        elif a > b and a + b == 8:  # 修改忘记改号的膜号
            for k in range(a - 1, 3, -1):
                d[i][1].append(d[i - 1][1][k])
                del d[i - 1][1][k]  # d[i - 1][1].remove(d[i - 1][1][k]) remove是对数值去除
        elif a == 8 and b == 1:  # 修改另一种忘记改号的情况
            for k in range(a - 1, 3, -1):
                d[i][1].append(d[i - 1][1][k])
                del d[i - 1][1][k]
            del d[i][1][0]

# 再进行膜号的分别输出
n = 1
m = 1
j = 0
for i in range(index + 1):
    if d[i][0] in wrong_pos:
        sheet2.write(n, 0, d[i][0])
        for j in range(len(d[i][1])):
            sheet2.write(n, j + 1, d[i][1][j], styleWrongPos)
        n += 1
    elif not d[i][0] in useless_pos and len(d[i][1]) != 4:
        sheet2.write(n, 0, d[i][0])
        for j in range(len(d[i][1])):
            sheet2.write(n, j + 1, d[i][1][j], styleMoreOrLess)
        n += 1
    elif not d[i][0] in useless_pos:
        sheet1.write(m, 0, d[i][0])
        for j in range(len(d[i][1])):
            sheet1.write(m, j + 1, d[i][1][j])
        tmp = [float(x) for x in d[i][1]]
        mean = np.mean(tmp)
        std = np.std(tmp)
        sheet1.write(m, j + 2, mean)
        sheet1.write(m, j + 3, std / mean)
        m += 1

# 找到并输出缺失膜号
monum = int(input('筛膜总张数')) + 1
exist = [False] * monum
n = 1
for data in d:
    if not data[0] in useless_pos and data[0] < monum:
        exist[data[0]] = True
for i in range(len(exist)):
    if not exist[i]:
        if i != 0:
            sheet2.write(n, 2, i)
            n += 1

# 导出工作簿
    wb.save('结果.xls')
