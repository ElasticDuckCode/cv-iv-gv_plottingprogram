
from program_libraries import *
from program_constants import *

'''
    This file contains plt_Coexist_Pressure function 
'''

def plt_Coexist_Pressure(xlsFile, plotDir, hminy, hmaxy):

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
