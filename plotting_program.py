#!/usr/bin/env python3

import os
import sys
import xlrd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates as dates
import datetime
from matplotlib.ticker import EngFormatter
from tkinter import filedialog
from tkinter import simpledialog
from tkinter import messagebox
from tkinter import *

'''
    Keithley MOScap Data Plotter

	28 August 2018
	Written by Jake Millhiser

    This program plots .xls spreadsheet data in a manner any real programmer
    would find distasteful. Includes copy and pasting my code rather than making a uniform function!

    Python3 Library Dependancies:
        xlrd
        numpy
        matplotlib
        tkinter

    Features to add:
        Selection of Specific Frequencies 
        Determine Lowest and Highest Frequencies using number of columns
'''
if os.path.exists('splash.txt'):
    with open('splash.txt','r') as fin:
        print(fin.read())

# Global Variables
windowTitle = 'XLS Plotter'
deviceArea = 16000e-8   # in cm^2
gline = 2   # graph line width

# Plot Flags
cvf = 0
gvf = 0
ivf = 0
hpf = 0

# Default C-V axis bounds
cminy = 0
#cmaxy = 3.5e-6
cmaxy = 1.5e-6

# Default G-V axis bounds
gminy = 1e4
gmaxy = 1e10

# Default I-V axis bounds
iminy = 1e-8
imaxy = 1e4

# Default HP axis bounds
hminy = 1e-6
hmaxy = 1e3
hminx = datetime.datetime(1,1,1,0,0,0)
hmaxx = datetime.datetime(1,1,1,1,0,0)

def cv_xlsplt(xlsFile):

    xlsBook = xlrd.open_workbook(xlsFile)
    xlsSheetNum = xlsBook.nsheets
    deviceNum =  xlsSheetNum - 2
    xlsSheetNames = xlsBook.sheet_names()
    print('Sheet names:\n\t',end='')
    for x in xlsSheetNames:
        print(x,end=' ')
    print('',end='\n')

    for i in range(xlsSheetNum):
        print('Plotting Sheet of index: ' + str(i))
        if i == 1 or i == 2:
            # do nothing, these are settings sheets
            print('\tSheets index 1 and 2 should have no data.')
        else:
            xlsSheet1 = xlsBook.sheet_by_index(i)
            deviceName = xlsSheetNames[i]

            # DC voltage bias value for x-axis
            x = [xlsSheet1.cell_value(j,2) for j in range(1,xlsSheet1.nrows)]

            # Capacitance value for y-axis
            y1 = [xlsSheet1.cell_value(j,0) for j in range(1,xlsSheet1.nrows)]
            y2 = [xlsSheet1.cell_value(j,4) for j in range(1,xlsSheet1.nrows)]
            y3 = [xlsSheet1.cell_value(j,8) for j in range(1,xlsSheet1.nrows)]
            y4 = [xlsSheet1.cell_value(j,12) for j in range(1,xlsSheet1.nrows)]
            y5 = [xlsSheet1.cell_value(j,16) for j in range(1,xlsSheet1.nrows)]
            y6 = [xlsSheet1.cell_value(j,20) for j in range(1,xlsSheet1.nrows)]
            y7 = [xlsSheet1.cell_value(j,24) for j in range(1,xlsSheet1.nrows)]
            y8 = [xlsSheet1.cell_value(j,28) for j in range(1,xlsSheet1.nrows)]
            y9 = [xlsSheet1.cell_value(j,32) for j in range(1,xlsSheet1.nrows)]
            y10 = [xlsSheet1.cell_value(j,36) for j in range(1,xlsSheet1.nrows)]
            y11 = [xlsSheet1.cell_value(j,40) for j in range(1,xlsSheet1.nrows)]
            y12 = [xlsSheet1.cell_value(j,44) for j in range(1,xlsSheet1.nrows)]
            y13 = [xlsSheet1.cell_value(j,48) for j in range(1,xlsSheet1.nrows)]
            y14 = [xlsSheet1.cell_value(j,52) for j in range(1,xlsSheet1.nrows)]
#            y15 = [xlsSheet1.cell_value(j,56) for j in range(1,xlsSheet1.nrows)]
#            y16 = [xlsSheet1.cell_value(j,60) for j in range(1,xlsSheet1.nrows)]
#            y17 = [xlsSheet1.cell_value(j,64) for j in range(1,xlsSheet1.nrows)]
#            y18 = [xlsSheet1.cell_value(j,68) for j in range(1,xlsSheet1.nrows)]
#            y19 = [xlsSheet1.cell_value(j,72) for j in range(1,xlsSheet1.nrows)]
#            y20 = [xlsSheet1.cell_value(j,76) for j in range(1,xlsSheet1.nrows)]
#            y21 = [xlsSheet1.cell_value(j,80) for j in range(1,xlsSheet1.nrows)]
#            y22 = [xlsSheet1.cell_value(j,84) for j in range(1,xlsSheet1.nrows)]
#            y23 = [xlsSheet1.cell_value(j,88) for j in range(1,xlsSheet1.nrows)]
#            y24 = [xlsSheet1.cell_value(j,92) for j in range(1,xlsSheet1.nrows)]
#            y25 = [xlsSheet1.cell_value(j,96) for j in range(1,xlsSheet1.nrows)]
#            y26 = [xlsSheet1.cell_value(j,100) for j in range(1,xlsSheet1.nrows)]
#            y27 = [xlsSheet1.cell_value(j,104) for j in range(1,xlsSheet1.nrows)]

            # Normalizing
            ny1 = [x/ deviceArea for x in y1]
            ny2 = [x/ deviceArea for x in y2]
            ny3 = [x/ deviceArea for x in y3]
            ny4 = [x/ deviceArea for x in y4]
            ny5 = [x/ deviceArea for x in y5]
            ny6 = [x/ deviceArea for x in y6]
            ny7 = [x/ deviceArea for x in y7]
            ny8 = [x/ deviceArea for x in y8]
            ny9 = [x/ deviceArea for x in y9]
            ny10 = [x/ deviceArea for x in y10]
            ny11 = [x/ deviceArea for x in y11]
            ny12 = [x/ deviceArea for x in y12]
            ny13 = [x/ deviceArea for x in y13]
            ny14 = [x/ deviceArea for x in y14]
#            ny15 = [x/ deviceArea for x in y15]
#            ny16 = [x/ deviceArea for x in y16]
#            ny17 = [x/ deviceArea for x in y17]
#            ny18 = [x/ deviceArea for x in y18]
#            ny19 = [x/ deviceArea for x in y19]


            # Graph
            fig = plt.figure(figsize=(6,7))
            cv = plt.subplot()

            # Normalized values
            cv.plot(x,ny1, linewidth=gline, label='10kHz')
            cv.plot(x,ny2, linewidth=gline, label='20kHz')
#            cv.plot(x,ny3, linewidth=gline, label='30kHz')
#            cv.plot(x,ny4, linewidth=gline, label='40kHz')
            cv.plot(x,ny5, linewidth=gline, label='50kHz')
#            cv.plot(x,ny6, linewidth=gline, label='60kHz')
#            cv.plot(x,ny7, linewidth=gline, label='70kHz')
#            cv.plot(x,ny8, linewidth=gline, label='80kHz')
#            cv.plot(x,ny9, linewidth=gline, label='90kHz')
            cv.plot(x,ny10, linewidth=gline, label='100kHz')
            cv.plot(x,ny11, linewidth=gline, label='200kHz')
#            cv.plot(x,ny12, linewidth=gline, label='300kHz')
#            cv.plot(x,ny13, linewidth=gline, label='400kHz')
            cv.plot(x,ny14, linewidth=gline, label='500kHz')
#            cv.plot(x,ny15, linewidth=gline, label='600kHz')
#            cv.plot(x,ny16, linewidth=gline, label='700kHz')
#            cv.plot(x,ny17, linewidth=gline, label='800kHz')
#            cv.plot(x,ny18, linewidth=gline, label='900kHz')
#            cv.plot(x,ny19, linewidth=gline, label='1MkHz')


            # Formatting
            cv.set_title(deviceName,loc='left',y=0) 
            engineeringFormat_Volts = EngFormatter(unit='V')
            engineeringFormat = EngFormatter()
            cv.xaxis.set_major_formatter(engineeringFormat)
            cv.yaxis.set_major_formatter(engineeringFormat)
            cv.set_xlabel('Gate Bias / (V)')
            cv.set_ylabel('Capacitance Density / (F/cm^2)')
            cv.set_xlim(-2,2)
            cv.set_ylim(ymin=cminy, ymax=cmaxy)

            cv.set_xticks(cv.get_xticks()[::2])
            #cv.set_yticks(cv.get_yticks()[::1])

            cv.legend(loc='upper right', fontsize='x-small')

            # Display Plot
            #plt.show()

            # Save and close Plot
            fig.savefig(plotDir + '/' + deviceName, dpi=300)
            plt.close(fig)

            # Decreasing device number
            deviceNum = deviceNum - 1 
    return

def gv_xlsplt(xlsFile):

    xlsBook = xlrd.open_workbook(xlsFile)
    xlsSheetNum = xlsBook.nsheets
    deviceNum =  xlsSheetNum - 2
    xlsSheetNames = xlsBook.sheet_names()
    print('Sheet names:\n\t',end='')
    for x in xlsSheetNames:
        print(x,end=' ')
    print('',end='\n')

    for i in range(xlsSheetNum):
        print('Plotting Sheet of index: ' + str(i))
        if i == 1 or i == 2:
            print(xlsSheetNames[i])
            # do nothing, these are settings sheets
            print('\tSheets index 1 and 2 should have no data.')
        else:
            xlsSheet1 = xlsBook.sheet_by_index(i)

            print(xlsSheetNames[i])
            deviceName = xlsSheetNames[i]

            # DC voltage bias value for x-axis
            x = [xlsSheet1.cell_value(j,2) for j in range(1,xlsSheet1.nrows)]

            # Capacitance value for y-axis
            y1 = [xlsSheet1.cell_value(j,1) for j in range(1,xlsSheet1.nrows)]
            y2 = [xlsSheet1.cell_value(j,5) for j in range(1,xlsSheet1.nrows)]
            y3 = [xlsSheet1.cell_value(j,9) for j in range(1,xlsSheet1.nrows)]
            y4 = [xlsSheet1.cell_value(j,13) for j in range(1,xlsSheet1.nrows)]
            y5 = [xlsSheet1.cell_value(j,17) for j in range(1,xlsSheet1.nrows)]
            y6 = [xlsSheet1.cell_value(j,21) for j in range(1,xlsSheet1.nrows)]
            y7 = [xlsSheet1.cell_value(j,25) for j in range(1,xlsSheet1.nrows)]
            y8 = [xlsSheet1.cell_value(j,29) for j in range(1,xlsSheet1.nrows)]
            y9 = [xlsSheet1.cell_value(j,33) for j in range(1,xlsSheet1.nrows)]
            y10 = [xlsSheet1.cell_value(j,37) for j in range(1,xlsSheet1.nrows)]
            y11 = [xlsSheet1.cell_value(j,41) for j in range(1,xlsSheet1.nrows)]
            y12 = [xlsSheet1.cell_value(j,45) for j in range(1,xlsSheet1.nrows)]
            y13 = [xlsSheet1.cell_value(j,49) for j in range(1,xlsSheet1.nrows)]
            y14 = [xlsSheet1.cell_value(j,53) for j in range(1,xlsSheet1.nrows)]
#            y15 = [xlsSheet1.cell_value(j,57) for j in range(1,xlsSheet1.nrows)]
#            y16 = [xlsSheet1.cell_value(j,61) for j in range(1,xlsSheet1.nrows)]
#            y17 = [xlsSheet1.cell_value(j,65) for j in range(1,xlsSheet1.nrows)]
#            y18 = [xlsSheet1.cell_value(j,69) for j in range(1,xlsSheet1.nrows)]
#            y19 = [xlsSheet1.cell_value(j,73) for j in range(1,xlsSheet1.nrows)]
#            y20 = [xlsSheet1.cell_value(j,77) for j in range(1,xlsSheet1.nrows)]
#            y21 = [xlsSheet1.cell_value(j,81) for j in range(1,xlsSheet1.nrows)]
#            y22 = [xlsSheet1.cell_value(j,85) for j in range(1,xlsSheet1.nrows)]
#            y23 = [xlsSheet1.cell_value(j,89) for j in range(1,xlsSheet1.nrows)]
#            y24 = [xlsSheet1.cell_value(j,93) for j in range(1,xlsSheet1.nrows)]
#            y25 = [xlsSheet1.cell_value(j,97) for j in range(1,xlsSheet1.nrows)]
#            y26 = [xlsSheet1.cell_value(j,101) for j in range(1,xlsSheet1.nrows)]
#            y27 = [xlsSheet1.cell_value(j,105) for j in range(1,xlsSheet1.nrows)]

            # Normalizing
            ny1 = [x/ deviceArea * 1e9 for x in y1]
            ny2 = [x/ deviceArea * 1e9 for x in y2]
            ny3 = [x/ deviceArea * 1e9 for x in y3]
            ny4 = [x/ deviceArea * 1e9 for x in y4]
            ny5 = [x/ deviceArea * 1e9 for x in y5]
            ny6 = [x/ deviceArea * 1e9 for x in y6]
            ny7 = [x/ deviceArea * 1e9 for x in y7]
            ny8 = [x/ deviceArea * 1e9 for x in y8]
            ny9 = [x/ deviceArea * 1e9 for x in y9]
            ny10 = [x/ deviceArea * 1e9 for x in y10]
            ny11 = [x/ deviceArea * 1e9 for x in y11]
            ny12 = [x/ deviceArea * 1e9 for x in y12]
            ny13 = [x/ deviceArea * 1e9 for x in y13]
            ny14 = [x/ deviceArea * 1e9 for x in y14]
#            ny15 = [x/ deviceArea * 1e9 for x in y15]
#            ny16 = [x/ deviceArea * 1e9 for x in y16]
#            ny17 = [x/ deviceArea * 1e9 for x in y17]
#            ny18 = [x/ deviceArea * 1e9 for x in y18]
#            ny19 = [x/ deviceArea * 1e9 for x in y19]


            # Graph
            fig = plt.figure(figsize=(6,7))
            gv = plt.subplot()

            # Normalized values
            gv.plot(x,ny1, linewidth=gline, label='10kHz')
            gv.plot(x,ny2, linewidth=gline, label='20kHz')
#            gv.plot(x,ny3, linewidth=gline, label='30kHz')
#            gv.plot(x,ny4, linewidth=gline, label='40kHz')
            gv.plot(x,ny5, linewidth=gline, label='50kHz')
#            gv.plot(x,ny6, linewidth=gline, label='60kHz')
#            gv.plot(x,ny7, linewidth=gline, label='70kHz')
#            gv.plot(x,ny8, linewidth=gline, label='80kHz')
#            gv.plot(x,ny9, linewidth=gline, label='90kHz')
            gv.plot(x,ny10, linewidth=gline, label='100kHz')
            gv.plot(x,ny11, linewidth=gline, label='200kHz')
#            gv.plot(x,ny12, linewidth=gline, label='300kHz')
#            gv.plot(x,ny13, linewidth=gline, label='400kHz')
            gv.plot(x,ny14, linewidth=gline, label='500kHz')
#            gv.plot(x,ny15, linewidth=gline, label='600kHz')
#            gv.plot(x,ny16, linewidth=gline, label='700kHz')
#            gv.plot(x,ny17, linewidth=gline, label='800kHz')
#            gv.plot(x,ny18, linewidth=gline, label='900kHz')
#            gv.plot(x,ny19, linewidth=gline, label='1MkHz')


            # Formatting
            gv.set_title(deviceName,loc='left',y=0) 
            engineeringFormat_Volts = EngFormatter(unit='V')
            engineeringFormat = EngFormatter()
            gv.xaxis.set_major_formatter(engineeringFormat)
            #gv.yaxis.set_major_formatter(engineeringFormat)
            gv.set_xlabel('Gate Bias / (V)')
            gv.set_ylabel('Conductance Density / (nS/cm^2)')
            gv.set_xlim(-2,2)
            gv.set_yscale('log')
            gv.set_ylim(ymin=gminy, ymax=gmaxy)

            gv.set_xticks(gv.get_xticks()[::2])
            #gv.set_yticks(gv.get_yticks()[::1])

            gv.legend(loc='upper right', fontsize='x-small')

            # Display Plot
            #plt.show()

            # Save and close Plot
            fig.savefig(plotDir + '/' + deviceName + '_GV', dpi=300)
            plt.close(fig)

            # Decreasing device number
            deviceNum = deviceNum - 1 
    return

def iv_xlsplt(xlsFile):

    xlsBook = xlrd.open_workbook(xlsFile)
    xlsSheetNum = xlsBook.nsheets
    deviceNum =  xlsSheetNum - 2
    xlsSheetNames = xlsBook.sheet_names()
    print('Sheet names:\n\t',end='')
    for x in xlsSheetNames:
        print(x,end=' ')
    print('',end='\n')

    for i in range(xlsSheetNum):
        print('Plotting Sheet of index: ' + str(i))
        if i == 1 or i == 2:
            # do nothing, these are settings sheets
            print('\tSheets of index 1 and 2 should have no data.')
        else:
            xlsSheet1 = xlsBook.sheet_by_index(i)
            deviceName = xlsSheetNames[i]

            # DC voltage bias value for x-axis
            x = [xlsSheet1.cell_value(j,1) for j in range(1,xlsSheet1.nrows)]

            # Current value for y-axis
            y1 = [xlsSheet1.cell_value(j,0) for j in range(1,xlsSheet1.nrows)]

            # Normalizing
            ny1 = [x/ deviceArea for x in y1]

            # Take Absolute Value for log scale
            ny1 = [abs(x) for x in ny1]

            # Graph
            fig = plt.figure(figsize=(6,7))
            iv = plt.subplot()

            # Normalized values
            iv.plot(x,ny1, linewidth=gline)


            # Formatting
            iv.set_title(deviceName,loc='left',y=0) 
            engineeringFormat_Volts = EngFormatter(unit='V')
            engineeringFormat = EngFormatter()
            iv.xaxis.set_major_formatter(engineeringFormat)
            #iv.yaxis.set_major_formatter(engineeringFormat)
            iv.set_xlabel('Gate Bias / (V)')
            iv.set_ylabel('Current Density / (A/cm^2)')
            iv.set_xlim(-2,2)
            iv.set_yscale('log')
            iv.set_ylim(ymin=iminy, ymax=imaxy)

            iv.set_xticks(iv.get_xticks()[::2])
            #iv.set_yticks(iv.get_yticks()[::1])

            # Display Plot
            #plt.show()

            # Save and close Plot
            fig.savefig(plotDir + '/' + deviceName + '_IV', dpi=300)
            plt.close(fig)

            # Decreasing device number
            deviceNum = deviceNum - 1 
    return

def plt_Coexist_Pressure(xlsFile):

    xlsBook = xlrd.open_workbook(xlsFile)
    xlsSheet1 = xlsBook.sheet_by_index(0)
    xlsSheetNames = xlsBook.sheet_names()
    deviceName = xlsSheetNames[0]
    print('Please be patient.')

    # Time Column
    x = [xlsSheet1.cell_value(j,0) for j in range(1,xlsSheet1.nrows)]

    # Pressure Columns
    y1 = [xlsSheet1.cell_value(j,1) for j in range(1,xlsSheet1.nrows)]
    y2 = [xlsSheet1.cell_value(j,2) for j in range(1,xlsSheet1.nrows)]
    y3 = [xlsSheet1.cell_value(j,3) for j in range(1,xlsSheet1.nrows)]
    y4 = [xlsSheet1.cell_value(j,4) for j in range(1,xlsSheet1.nrows)]
    y5 = [xlsSheet1.cell_value(j,5) for j in range(1,xlsSheet1.nrows)]
    y6 = [xlsSheet1.cell_value(j,6) for j in range(1,xlsSheet1.nrows)]

    # Graph
    fig = plt.figure(figsize=(8,8))
    hp = plt.subplot()
    hp.plot(x,y1, linewidth=gline, label='LWF Chamber')
    hp.plot(x,y2, linewidth=gline, label='Load Lock 2')
    hp.plot(x,y3, linewidth=gline, label='Oxide Chamber')
    hp.plot(x,y4, linewidth=gline, label='Plasma Chamber')
    hp.plot(x,y5, linewidth=gline, label='Metal Chamber')
    hp.plot(x,y6, linewidth=gline, label='Loac Lock 2')

    # Formatting
    hp.set_title('CoExist Pressure Readings ' + deviceName ,loc='left',y=0) 
    engineeringFormat_Volts = EngFormatter(unit='V')
    engineeringFormat = EngFormatter()
    #hp.xaxis.set_major_formatter(engineeringFormat)
    hp.yaxis.set_major_formatter(engineeringFormat)
    hp.grid()
    hp.set_xlabel('Time')
    hp.set_ylabel('Pressure (in Torr)')
    #hp.set_xlim(dates.date2num([hminx,hmaxx]))
    hp.set_ylim(hminy, hmaxy)
    hp.set_xticks(hp.get_xticks()[::100])
    #print(hp.get_xticks()[::100])
    #hp.set_yticks(hp.get_yticks()[::1])
    hp.set_yscale('log')
    #plt.xticks(rotation=90)
    plt.locator_params(axis='x',nbins=15)
    plt.gcf().autofmt_xdate()
    hp.legend(loc='upper right', fontsize='x-small')

    # Save and close Plot
    fig.savefig(plotDir + '/' + deviceName + '_HP', dpi=300)
    plt.close(fig)

    return

def exitProgram():
    sys.exit(0)

def printFlag():
    print(cv_flag.get())

def setCV():
    global cvf
    cvf = 1
    print(cvf)

def setGV():
    global gvf
    gvf = 1
    print(gvf)

def setIV():
    global ivf
    ivf = 1
    print(ivf)

def setHP():
    global hpf
    hpf = 1
    print(hpf)

def unsetCV():
    global cvf
    cvf = 0
    print(cvf)

def unsetGV():
    global gvf
    gvf = 0
    print(gvf)

def unsetIV():
    global ivf
    ivf = 0
    print(ivf)

def unsetHP():
    global hpf
    hpf = 0
    print(hpf)

# Create Exit Prompt
root = Tk()
root.withdraw()
root.title(windowTitle)
quitPrompt = Label(root, text="Would you like to plot another .xls file?")
b = Button(root, text="Yes", command=root.quit)
c = Button(root, text="No", command=exitProgram)
quitPrompt.pack()
b.pack()
c.pack()

# Create Plot Type Selection
plotSel = Tk()
plotSel.withdraw()
plotSel.title(windowTitle)
Label(plotSel, text="Select graph types you wish to plot").pack()
Label(plotSel, text="").pack()
Label(plotSel, text="C-V Plot").pack()
Button(plotSel, text="Enable", command=setCV, pady=5).pack()
Button(plotSel, text="Disable", command=unsetCV, pady=5).pack()
Label(plotSel, text="").pack()
Label(plotSel, text="G-V Plot").pack()
Button(plotSel, text="Enable", command=setGV, pady=5).pack()
Button(plotSel, text="Disable", command=unsetGV, pady=5).pack()
Label(plotSel, text="").pack()
Label(plotSel, text="I-V Plot").pack()
Button(plotSel, text="Enable", command=setIV, pady=5).pack()
Button(plotSel, text="Disable", command=unsetIV, pady=5).pack()
Label(plotSel, text="").pack()
Label(plotSel, text="Harshil's Pressure Plot").pack()
Button(plotSel, text="Enable", command=setHP).pack()
Button(plotSel, text="Disable", command=unsetHP).pack()
Label(plotSel, text="").pack()
Button(plotSel, text="Ok", command=plotSel.quit).pack()

# Get CV Boundaries Window
boundCV = Tk()
boundCV.withdraw()
boundCV.title(windowTitle)
Label(boundCV, text="Minimum Capacitance").grid(row=0)
Label(boundCV, text="Maximum Capacitance").grid(row=1)
e1 = Entry(boundCV)
e2 = Entry(boundCV)
e1.grid(row=0,column=1)
e2.grid(row=1,column=1)
e1.insert(10,str(cminy))
e2.insert(10,str(cmaxy))
Button(boundCV, text='Ok', command=boundCV.quit).grid(row=2)

# Get GV Boundaries Window
boundGV = Tk()
boundGV.withdraw()
boundGV.title(windowTitle)
Label(boundGV, text="Minimum Conductance").grid(row=0)
Label(boundGV, text="Maximum Conductance").grid(row=1)
g1 = Entry(boundGV)
g2 = Entry(boundGV)
g1.grid(row=0,column=1)
g2.grid(row=1,column=1)
g1.insert(10,str(gminy))
g2.insert(10,str(gmaxy))
Button(boundGV, text='Ok', command=boundGV.quit).grid(row=2)

# Get IV Boundaries Window
boundIV = Tk()
boundIV.withdraw()
boundIV.title(windowTitle)
Label(boundIV, text="Minimum Current").grid(row=0)
Label(boundIV, text="Maximum Current").grid(row=1)
i1 = Entry(boundIV)
i2 = Entry(boundIV)
i1.grid(row=0,column=1)
i2.grid(row=1,column=1)
i1.insert(10,str(iminy))
i2.insert(10,str(imaxy))
Button(boundIV, text='Ok', command=boundIV.quit).grid(row=2)

# Get HP Boundaries Window
boundHP = Tk()
boundHP.withdraw()
boundHP.title(windowTitle)
Label(boundHP, text="Minimum Pressure").grid(row=0)
Label(boundHP, text="Maximum Pressure").grid(row=1)
h1 = Entry(boundHP)
h2 = Entry(boundHP)
h1.grid(row=0,column=1)
h2.grid(row=1,column=1)
h1.insert(10,str(hminy))
h2.insert(10,str(hmaxy))
Button(boundHP, text='Ok', command=boundHP.quit).grid(row=2)

# Display First Run Warning
messagebox.showinfo(windowTitle, "Make sure your .xls sheets are named after the sample you are plotting. The sheet name will appear on the graph.")

while True:
    root.withdraw()
    xlsFile = filedialog.askopenfilename(title="Select .xls file to open", filetypes = (("xls files", "*.xls"),("xlsx files", "*.xlsx")))
    #addedTags = simpledialog.askstring(windowTitle, 'Give additional tags (FGA, FGA2, etc.)')
    plotDir = filedialog.askdirectory(title = "Choose Directory to Save Plots")
    plotSel.deiconify()
    plotSel.mainloop()
    plotSel.withdraw()
    if cvf == 1:
        boundCV.deiconify()
        boundCV.mainloop()
        cminy = float(e1.get())
        cmaxy = float(e2.get())
        boundCV.withdraw()
        cv_xlsplt(xlsFile)
    if gvf == 1:
        boundGV.deiconify()
        boundGV.mainloop()
        gminy = float(g1.get())
        gmaxy = float(g2.get())
        boundGV.withdraw()
        gv_xlsplt(xlsFile)
    if ivf == 1:
        boundIV.deiconify()
        boundIV.mainloop()
        iminy = float(i1.get())
        imaxy = float(i2.get())
        boundIV.withdraw()
        iv_xlsplt(xlsFile)
    if hpf == 1:
        messagebox.showinfo(windowTitle, "Please include only times you wish to plot. This program does not yet have user-given time bounds for this mode yet.")
        boundHP.deiconify()
        boundHP.mainloop()
        hminy = float(h1.get())
        hmaxy = float(h2.get())
        boundHP.withdraw()
        plt_Coexist_Pressure(xlsFile)
    root.deiconify()
    root.mainloop()
