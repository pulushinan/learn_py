""" 0.1.2更新日志
    添加了修改因为录入错误和未及时改号等而导致膜号错误的功能"""

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

# 将数据写入预定单元格
n = 1
m = 1
j = 0
sheet1.write(0, 0, '膜号')
sheet1.write(0, 1, 'RI值')
sheet1.write(0, 2, 'RI值')
sheet1.write(0, 3, 'RI值')
sheet1.write(0, 4, 'RI值')
sheet1.write(0, 5, 'AVG')
sheet1.write(0, 6, 'CV')
sheet2.write(0, 0, '膜号')
sheet2.write(0, 1, 'RI值')
for i in range(index + 1):
    if d[i][0] in wrong_pos:
        sheet2.write(n, 0, d[i][0])
        for j in range(len(d[i][1])):
            sheet2.write(n, j + 1, d[i][1][j], styleWrongPos)
        n += 1
    elif len(d[i][1]) != 4:
        sheet2.write(n, 0, d[i][0])
        for j in range(len(d[i][1])):
            sheet2.write(n, j + 1, d[i][1][j], styleMoreOrLess)
        n += 1
    else:
        sheet1.write(m, 0, d[i][0])
        for j in range(len(d[i][1])):
            sheet1.write(m, j + 1, d[i][1][j])
        tmp = [float(x) for x in d[i][1]]  # 这种句式不懂
        mean = np.mean(tmp)
        std = np.std(tmp)
        sheet1.write(m, j + 2, mean)
        sheet1.write(m, j + 3, std / mean)
        m += 1

# 导出工作簿
name = input('请输入文件名称\n叫 筛膜数据处理结果 可以么？\n请输入 \"可以\" 或者 文件名称\n')
if name == '可以':
    wb.save('筛膜数据处理结果.xls')
else:
    wb.save(name + '.xls')

# 结束提示
print('请查看文件')
