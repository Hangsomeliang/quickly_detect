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

#airPLS迭代自适应加权惩罚最小二乘法（去除背景）
def airPLS(X, lam, order, wep=0.05, p=0.05,itermax=20):
    (m,n) = X.shape
    Z = np.empty([m,n])
    D = np.diff(np.eye(n), order, axis=0)#计算离散差分（Zi - Zi-1）
    DD = np.matmul(lam * D.T,D)#离散差分求和lam*D'*D
    for i in range(m):
        w = np.ones([n,1]).T#生产全为1的一维数组
        x = X[i,:]#取第i行数据
        for j in range(1,itermax+1):
            W = spdiags(w, 0, n, n)#按主对角线排布w
            C = cholesky(W + DD)
            z = np.matmul(np.linalg.inv(C), np.matmul(np.linalg.inv(C.T), (w * x).T)).T
            d = x - z
            dssn = np.abs(sum(d[d<0]))
            if dssn < 0.001 * sum(np.abs(x)):
                break
            w[d>=0] = 0
            w[0][:math.ceil(n*wep)] = p
            w[0][n-math.floor(n*wep)-1:] = p
            to_exp = abs(d[d<0])/dssn
            w[d<0] = j * np.exp(to_exp)
        Z[i,:] = z
    Xc = X-Z
    return Xc,Z