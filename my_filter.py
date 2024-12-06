from openpyxl import load_workbook
import numpy as np
import matplotlib.pyplot as plt
from scipy.datasets import electrocardiogram
from scipy.signal import find_peaks
from scipy.sparse import spdiags
from scipy.linalg import cholesky
from scipy.signal import savgol_filter
import math
import tkinter as tk
from tkinter import filedialog
import openpyxl
import os
import pywt
from sympy import *

def myfilter(data):
    print(len(data))
    wavelet_basis = 'db8'
    w = pywt.Wavelet(wavelet_basis)
    coeffs = pywt.wavedec(data,wavelet_basis,level=3)
    for j in range(1,len(coeffs)):
        tmp = coeffs[j].copy()
        sum = 0.0
        for z in coeffs[j]:
            sum = sum + abs(z)
        N = len(coeffs[j])
        sum = (1.0/float(N))* sum
        sigma = (1/0.6745)*sum
        threshold = sigma * math.sqrt(2.0 * math.log(float(N),math.e))
        coeffs[j] = pywt.threshold(coeffs[j],threshold)
    signalrec = pywt.waverec(coeffs,wavelet_basis)
    signalrec = np.array(signalrec)
    return signalrec
