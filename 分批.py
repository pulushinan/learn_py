# 载入库
import xlrd
import xlwt
from xlwt import Workbook
import numpy as np


# 定义偏差函数
def dev(x, y):
    n = x - y
    n = n / x
    return n
