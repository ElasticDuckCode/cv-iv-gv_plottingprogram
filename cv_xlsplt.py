
from program_libraries import *
from program_constants import *

# the above does not seem to work to define dit_calc
from dit_calc import *

'''
    This file contains cv_xlsplt function 
'''

def cv_xlsplt(xlsFile, plotDir, cminy, cmaxy, display_dit=0):

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

            if display_dit == 1:
                Dit_Display = dit_calc(xlsFile, plotDir, i, 1)
                # TEMPORARILY DISABLED UNTIL THRESHOLD AND FLATBAND VOLTAGES MADE cv.text(0,cmaxy - (cmaxy/20), u"D\u1D62\u209C = " + "{:.6e}".format(Dit_Display[0][1]), horizontalalignment='center', fontsize=ditsize, fontweight='bold') # place text 5% down from top, also u means unicode so there can be subscript

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
