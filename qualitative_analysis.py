from openpyxl import load_workbook
import numpy as np
import matplotlib.pyplot as plt
from scipy.datasets import electrocardiogram
from scipy.signal import find_peaks
from scipy.sparse import spdiags
from scipy.linalg import cholesky
import math
import tkinter as tk
from tkinter import filedialog
import openpyxl
import os
from Binary_search import binary_search

#打开特征X射线数据库（XRAYLIB）
xraylib_book = load_workbook("选峰表.xlsx")
#打开6种不同检测参数的数据库
##1号参数
xraylib_sheet1 = xraylib_book['1']
xraylib1_theta = []
xraylib1_element = []
for row in xraylib_sheet1.iter_rows():
    theta = row[9].value
    element = row[1].value
    if theta == None:
        continue
    xraylib1_theta.append(theta)
    xraylib1_element.append(element)
##2号参数
xraylib_sheet2 = xraylib_book['2']
xraylib2_theta = []
xraylib2_element = []
for row in xraylib_sheet2.iter_rows():
    theta = row[8].value
    element = row[1].value
    if theta == None:
        continue
    xraylib2_theta.append(theta)
    xraylib2_element.append(element)
##3号参数
xraylib_sheet3 = xraylib_book['3']
xraylib3_theta = []
xraylib3_element = []
for row in xraylib_sheet3.iter_rows():
    theta = row[8].value
    element = row[1].value
    if theta == None:
        continue
    xraylib3_theta.append(theta)
    xraylib3_element.append(element)
##4号参数
xraylib_sheet4 = xraylib_book['4']
xraylib4_theta = []
xraylib4_element = []
for row in xraylib_sheet4.iter_rows():
    theta = row[8].value
    element = row[1].value
    if theta == None:
        continue
    xraylib4_theta.append(theta)
    xraylib4_element.append(element)
##5号参数
xraylib_sheet5 = xraylib_book['5']
xraylib5_theta = []
xraylib5_element = []
for row in xraylib_sheet5.iter_rows():
    theta = row[8].value
    element = row[1].value
    if theta == None:
        continue
    xraylib5_theta.append(theta)
    xraylib5_element.append(element)
##6号参数
xraylib_sheet6 = xraylib_book['6']
xraylib6_theta = []
xraylib6_element = []
for row in xraylib_sheet6.iter_rows():
    theta = row[8].value
    element = row[1].value
    if theta == None:
        continue
    xraylib6_theta.append(theta)
    xraylib6_element.append(element)
#选择并打开检测源文件

