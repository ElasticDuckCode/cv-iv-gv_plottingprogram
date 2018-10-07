
from program_libraries import *
from program_constants import *

'''
    This file contains gv_xlsplt function 
'''

def gv_xlsplt(xlsFile, plotDir, gminy, gmaxy):

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
