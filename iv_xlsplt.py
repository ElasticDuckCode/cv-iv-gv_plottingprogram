
from program_libraries import *
from program_constants import *

'''
    This file contains iv_xlsplt function 
'''

def iv_xlsplt(xlsFile, plotDir, iminy, imaxy):

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
