
from program_libraries import *
from program_constants import *

'''
    This file contains dit_calc function 
'''

def dit_calc(xlsFile, plotDir, index=0, cv_plt_enable=0):

    xlsBook = xlrd.open_workbook(xlsFile)
    xlsSheetNum = xlsBook.nsheets
    deviceNum =  xlsSheetNum - 2
    xlsSheetNames = xlsBook.sheet_names()
    print('Sheet names:\n\t',end='')
    for x in xlsSheetNames:
        print(x,end=' ')
    print('',end='\n')

    Dit_Vector = []
    for i in range(xlsSheetNum):
        if cv_plt_enable == 1:
            i = index
        print('Plotting Sheet of index: ' + str(i))
        if i == 1 or i == 2:
            print('\tSheets index 1 and 2 should contain no measurement data.')
        else:
            xlsSheet = xlsBook.sheet_by_index(i)
            deviceName = xlsSheetNames[i]
            
            Gpw = [] # python array, not numpy array!
            Vdc = []
            # Getting all data as a matrix
            for row in range(xlsSheet.nrows):
                _Gpw = []
                if row == 0:
                    continue  # first row is just labels

                for col in range(0,xlsSheet.ncols,4):
                    # Columns go as [C G V_DC f] for each sweep
                    # Columns go as [0 1  2   3] for each sweep
                    # col will be an pointer to C for every measuremetn, thus G is col + 1, etc
		    # numerator = omega * G * C_ox^2
		    # denominator = G^2 + omega^2(C_ox - C)^2
                    numerator = (2 * np.pi) * xlsSheet.cell_value(row,col+3) * xlsSheet.cell_value(row,col+1) * xlsSheet.cell_value(xlsSheet.nrows-4,col)**2

                    denominator = xlsSheet.cell_value(row,col+1)**2 + 4 * np.pi**2 * xlsSheet.cell_value(row,col+3)**2 * (xlsSheet.cell_value(xlsSheet.nrows-4,col) - xlsSheet.cell_value(row,col))**2
                    _Gpw.append(numerator / denominator)

                Gpw.append(_Gpw)
                Vdc.append(xlsSheet.cell_value(row,2))

            Dit = []
            for row in range(xlsSheet.nrows-1): # subtract 1 since we skipped the row of labels
                _Gpw = Gpw[row]
                # Converts python array to numpy array, then takes maximum value
                Dit.append(np.amax(np.asarray(_Gpw)) * (2.5 / (deviceArea * e)))

            Dit_withBias = np.stack((Vdc,Dit), axis=-1) # -1 means append as cols not rows
            np.savetxt(plotDir + '/' + deviceName + "_DIT" + ".csv", Dit_withBias, delimiter=',')


            # Select max Dit only from ~0 to ~1V bias
            numRows = np.shape(Dit_withBias)
            numRows = numRows[0]
            if Vdc[0] - Vdc[1] > 0:
                # Sweep goes from positive to negative
                Max_Dit = Dit_withBias[:-int(numRows/3),:] # remove last half of data
                Max_Dit = Max_Dit[int(numRows/4):,:] # then remove first half of half

            else:
                Max_Dit = Dit_withBias[int(numRows/3):,:] # remove first half of data
                Max_Dit = Max_Dit[:-int(numRows/4),:] # then remove last half of half

            #DEBUG_PURPOSES np.savetxt(plotDir + '/' + deviceName + "_DIT_Reduced" + ".csv", Max_Dit, delimiter=',')

            Max_Dit = np.amax(Max_Dit[:,1])
            if cv_plt_enable == 0:
                f = open(plotDir + '/' + deviceName + "_DIT" + ".txt", 'w')
                f.write(deviceName + '\n\tD_it = ' + "{:.5e}".format(Max_Dit))
                f.close()
            Dit_Vector.append([deviceName,Max_Dit])
        print(Dit_Vector)
        if cv_plt_enable == 1:
            break
    return Dit_Vector




#            _data = []
#            for row in range (xlsSheet.ncols):
#                _row = []
#                for col in range (xlsSheet.ncols):
#                    _row.append(xlsSheet.cell_value(row,col))
#                _data.append(_row)


