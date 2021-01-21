# -*- coding: utf-8 -*-
"""
Created on Sun Jan  3 19:36:38 2021

@author: Constant
"""

import openpyxl
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import random

def probability(df):
    s = np.sum(df, axis=0)
    m = len(df)
    mu = s/m
    vr = np.sum((df - mu)**2, axis=0)
    variance = vr/m
    var_dia = np.diag(variance)
    k = len(mu)
    X = df - mu
    ps = 1/((2*np.pi)**(k/2)*(np.linalg.det(var_dia)**0.5))* np.exp(-0.5* np.sum(X @ np.linalg.pinv(var_dia) * X,axis=1))
    p = np.array(ps.values.tolist())
    return p


def tpfpfn(ep, p):
    tp, fp, fn = 0, 0, 0
    for i in range(len(y)):

        if p[i] <= ep and y[i][0] == 1:

            tp += 1
        elif p[i] <= ep and y[i][0] == 0:

            fp += 1
        elif p[i] > ep and y[i][0] == 1:

            fn += 1
    return tp, fp, fn



def f1(ep, p):
    tp, fp, fn = tpfpfn(ep,p)
    prec = tp/(tp + fp)
    rec = tp/(tp + fn)
    f1 = 2*prec*rec/(prec + rec)
    return f1


while True:
    choice= input("Traiter les données [y/n] ?") 
    if choice == "y":
        wb = openpyxl.load_workbook('Data.xlsx')
        sheet = wb['Data']
        sheet.delete_cols(1,2)

        for i in range(50):
            cellx = 'A' + str(i+2)
            celly = 'B' + str(i+2)
            x = random.gauss(5,1.5)
            y = random.gauss(5,1.5)
            sheet[cellx].value = x
            sheet[celly].value = y
        wb.save('Data.xlsx')
        
        df = pd.read_excel('Data.xlsx', sheet_name='Data', header=None)
        cvx = pd.read_excel('Data.xlsx', sheet_name='Cross-validation', header=None)
        cvy = pd.read_excel('Data.xlsx', sheet_name='Label', header=None)
        df = df.drop([0])
        cvx = cvx.drop([0])
        cvy = cvy.drop([0])
        y = np.array(cvy)
        
        
        plt.figure()
        plt.scatter(df[0], df[1])
        plt.axis([0,10,0,10])
        plt.show()
        
        
        p = probability(df)
        p1 = probability(cvx)
        eps = [i for i in p1 if i <= p1.mean()]
        
        f = []
        for i in eps:
            f.append(f1(i, p1))
            
        
        idx = np.array(f).argmax()
        
        e = eps[idx]
        
        normal = df.copy()
        anomaly = df.copy()
        label = []
        
        wb = openpyxl.load_workbook('Data.xlsx')
        sheet = wb['Data']
        sheet2 = wb['Cross-validation']
        sheet3 = wb['Label']
        
        
        
        for i in range(len(df)):
            cellx = 'A' + str(i+2+len(y))
            celly = 'B' + str(i+2+len(y))
            sheet2[cellx].value = df.iloc[i,0]
            sheet2[celly].value = df.iloc[i,0]
        
            if p[i] <= e:
                label.append(1)
                normal = normal.drop([i+1])
                sheet3[cellx].value = 1
            else:
                label.append(0)
                anomaly = anomaly.drop([i+1])
                sheet3[cellx].value = 0
        
        plt.figure()
        plt.scatter(normal[0], normal[1], color = 'black')
        plt.scatter(anomaly[0], anomaly[1], color = 'red')
        plt.axis([0,10,0,10])
        plt.show()
        
        
        wb.save('Data.xlsx')
        
    else:
        break
