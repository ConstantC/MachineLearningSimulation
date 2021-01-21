# -*- coding: utf-8 -*-
"""
Created on Sun Jan  3 19:36:38 2021

@author: Constant
"""
import openpyxl
import random

wb = openpyxl.load_workbook('Data.xlsx')
sheet2 = wb['Cross-validation']
sheet3 = wb['Label']
sheet2.delete_cols(1,2)
sheet3.delete_cols(1,2)

for i in range(150):
    cellx = 'A' + str(i+2)
    celly1 = 'B' + str(i+2)
    xcv = random.gauss(5,2)
    ycv = random.gauss(5,2)
    
    if 4 <= xcv <= 6 and 4 <= xcv <= 6:
        yl = 0
    else:
        yl = 1
    if i >= 130:
        idx = random.randint(0,150)
        yl = random.randint(0,1)
        
    sheet2[cellx].value = xcv
    sheet2[celly1].value = ycv
    sheet3[cellx].value = yl

wb.save('Data.xlsx')