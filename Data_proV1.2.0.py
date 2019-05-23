# 载入库
import numpy as np
import xlrd
import xlwt
from xlwt import Workbook

# 交互界面
print('请稍等...\n')
namein = input('请输入原始数据文件名称\n')
el = float(input('初筛还是复筛？\n初筛请输入1，复筛请输入2\n'))

if el == 1:
    monum = int(input('请输入筛膜总张数\n')) + 1

downT = float(input('T带位置的下限是?\n'))
upT = float(input('T带位置的上限是?\n'))
absT = float(input('T带值的下限是?\n'))
absC = float(input('C带值的下限是?\n'))

if el == 1:
    col = int(input('请输入初筛膜号所在列数，请用数字表示\n')) - 1
elif el == 2:
    col = int(input('请输入f复筛膜号所在列数，请用数字表示\n')) - 1

# 载入原始Excel表
loc = (namein + '.xlsx')

# 打开表格及工作簿
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
if el == 2:
    sheet_re = wb.sheet_by_index(1)

# 行数
row = sheet.nrows
if el == 2:
    row_re = sheet_re.nrows

# 载入膜号列数
col_of_index = [col]
if el == 2:
    col_re_of_index = [col]
# 数据处理、转置、编组、筛选出不合格膜号
# 这里仅仅筛选出位置前移和绝对值偏低的
d = []
locfailed = set()
absfailed = set()
index = -1
er = 0
for j in col_of_index:
    for i in range(row):
        cur = sheet.cell_value(i, j)
        if type(cur) == float:
            T = float(sheet.cell_value(i, j + 1))
            C = float(sheet.cell_value(i, j + 2))
            t = float(sheet.cell_value(i, j + 3))
            c = float(sheet.cell_value(i, j + 4))
            rValue = sheet.cell_value(i, j + 5)
            er = 0
            if T <= downT or T >= upT:  # 筛选T位置偏移
                er += 1
                if er >= el:
                    locfailed.add(int(cur))
            if t <= absT:  # 筛选T绝对值低
                er += 1
                if er >= el:
                    absfailed.add(int(cur))
            if c <= absC:  # 筛选C绝对值低
                er += 1
                if er >= el:
                    absfailed.add(int(cur))
            if type(rValue) == float or type(rValue) == int:
                if index < 0 or d[index][0] != int(cur):
                    d.append((int(cur), [rValue]))
                    index += 1
                else:
                    d[index][1].append(rValue)

# 出错膜号
s = set()
wrong_pos = set()
for i in range(index + 1):
    if d[i][0] in s:
        wrong_pos.add(d[i][0])
    else:
        s.add(d[i][0])

if el == 1:
    # 将处理好的数据保存
    # 创建工作簿
    wb = Workbook()
    sheet1 = wb.add_sheet('数据处理', cell_overwrite_ok=True)
    sheet2 = wb.add_sheet('出错数据', cell_overwrite_ok=True)

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
    sheet2.write(0, 0, '重错膜号')
    sheet2.write(0, 2, '缺失膜号')
    sheet2.write(0, 4, '位置偏移')
    sheet2.write(0, 6, '绝对值偏低')
    sheet2.write(0, 8, 'CV不合格')
    sheet2.write(0, 10, '复筛膜号汇总')

# 先进行错误膜号的修改
useless_pos = set()  # 输出错号,修改录入出错的情况再进行录入
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
    if d[i][0] > monum:
        useless_pos.add(d[i][0])
# 再进行膜号的分别输出
n = 1
m = 1
j = 0
less_pos = set()
cvfailed = set()
for i in range(index + 1):
    if d[i][0] in wrong_pos:
        if el == 1:
            sheet2.write(n, 0, d[i][0], styleWrongPos)
            n += 1
    elif not d[i][0] in useless_pos and len(d[i][1]) != 4:
        less_pos.add(d[i][0])
        if el == 1:
            sheet2.write(n, 0, d[i][0], styleMoreOrLess)
            n += 1
    elif not d[i][0] in useless_pos and not d[i][0] in locfailed and not d[i][0] in absfailed:
        tmp = [float(x) for x in d[i][1]]
        mean = np.mean(tmp)
        std = np.std(tmp)  # 对样本数据求标准差
        cv = std / mean
        if cv <= 0.15:
            if el == 1:
                sheet1.write(m, 0, d[i][0])
                for j in range(len(d[i][1])):
                    sheet1.write(m, j + 1, d[i][1][j])
                    sheet1.write(m, j + 2, mean)
                    sheet1.write(m, j + 3, std / mean)
                m += 1
        else:
            cvfailed.add(d[i][0])

# 找到并输出缺失膜号
lose_pos = set()
exist = [False] * monum
n = 1
for data in d:
    if not data[0] in useless_pos and data[0] < monum:
        exist[data[0]] = True
for i in range(len(exist)):
    if not exist[i]:
        if i != 0:
            lose_pos.add(i)
            if el == 1:
                sheet2.write(n, 2, i)
                n += 1

# 输出不合格膜号
n = 1
for i in locfailed:
    if i not in useless_pos and i not in absfailed and i not in cvfailed:
        if el == 1:
            sheet2.write(n, 4, i)
            n += 1
n = 1
for i in absfailed:
    if i not in useless_pos and i not in cvfailed:
        if el == 1:
            sheet2.write(n, 6, i)
            n += 1
n = 1
for i in cvfailed:
    if i not in useless_pos:
        if el == 1:
            sheet2.write(n, 8, i)
            n += 1
allfailed = set()
allfailed = locfailed | absfailed | cvfailed | wrong_pos | lose_pos | less_pos
n = 1
x = 0
L = (len(allfailed) // 10)
failed = list(allfailed)
failed.sort()
if el == 1:
    for i in allfailed:
        if i not in useless_pos:
            if n <= L:
                sheet2.write(n, 10 + x, i)
                n += 1
            elif n > L:
                n = 1
                x += 1
                sheet2.write(n, 10 + x, i)
                n += 1

# 对复筛数据的处理


# 导出工作簿
name = input('请输入文件名称\n叫 筛膜数据处理结果 可以么？\n请输入 \"可以\" 或者 文件名称\n')
if name == '可以':
    wb.save('筛膜数据处理结果.xls')
else:
    wb.save(name + '.xls')

# 结束提示
print('请查看文件')
