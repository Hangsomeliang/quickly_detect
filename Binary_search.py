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

#二分查找函数定义
def binary_search(arr, target):
    if not arr:
        return None
    # arr.sort()#数组排序

    low = 0
    high = len(arr) - 1

    while low <= high:
        mid = (low+high)//2
        if arr[mid] == target:
            return mid
        elif arr[mid] < target:
            low = mid + 1
        else:
            high = mid - 1
    if low >= len(arr):
        return high
    if abs(target - arr[low]) < abs(target-arr[high]):
        return low
    else:
        return high