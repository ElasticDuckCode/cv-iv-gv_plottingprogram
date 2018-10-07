#!/usr/bin/env python3

from program_libraries import *
from program_constants import *

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
# Plot Flags
cvf = 0
gvf = 0
ivf = 0
hpf = 0
ditf = 0

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

if os.path.exists('splash.txt'):
    with open('splash.txt','r') as fin:
        print(fin.read())

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

def setDIT():
    global ditf
    ditf = 1
    print(ditf)

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

def unsetDIT():
    global ditf
    ditf = 0
    print(ditf)

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
Label(plotSel, text="Select graph types you wish to plot:").grid(row=0,column=0,pady=10)
Label(plotSel, text="Capacitance-Voltage (CV)").grid(row=1, column=0)
Button(plotSel, text="Enable", command=setCV).grid(row=1,column=1)
Button(plotSel, text="Disable", command=unsetCV).grid(row=1,column=2)
Label(plotSel, text="Conductance-Voltage (GV)").grid(row=2,column=0)
Button(plotSel, text="Enable", command=setGV).grid(row=2,column=1)
Button(plotSel, text="Disable", command=unsetGV).grid(row=2,column=2)
Label(plotSel, text="DIT Calculation (No Plot)").grid(row=3,column=0)
Button(plotSel, text="Enable", command=setDIT).grid(row=3,column=1)
Button(plotSel, text="Disable", command=unsetDIT).grid(row=3,column=2)
Label(plotSel, text="Current-Voltage (IV)").grid(row=4,column=0)
Button(plotSel, text="Enable", command=setIV).grid(row=4,column=1)
Button(plotSel, text="Disable", command=unsetIV).grid(row=4,column=2)
Label(plotSel, text="Harshil's Pressure Data").grid(row=5,column=0)
Button(plotSel, text="Enable", command=setHP).grid(row=5,column=1)
Button(plotSel, text="Disable", command=unsetHP).grid(row=5,column=2)
Button(plotSel, text="Ok", command=plotSel.quit).grid(row=6,column=0)

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
    if cvf == 1 and ditf == 1:
        #messagebox.showinfo(windowTitle, "With both CV and DIT enabled, the CV plots will display the DIT value.")
        boundCV.deiconify()
        boundCV.mainloop()
        cminy = float(e1.get())
        cmaxy = float(e2.get())
        boundCV.withdraw()
        cv_xlsplt(xlsFile, plotDir, cminy, cmaxy, 1)
    elif cvf == 1:
        boundCV.deiconify()
        boundCV.mainloop()
        cminy = float(e1.get())
        cmaxy = float(e2.get())
        boundCV.withdraw()
        cv_xlsplt(xlsFile, plotDir, cminy, cmaxy)
    elif ditf == 1:
        dit_calc(xlsFile,plotDir)
    if gvf == 1:
        boundGV.deiconify()
        boundGV.mainloop()
        gminy = float(g1.get())
        gmaxy = float(g2.get())
        boundGV.withdraw()
        gv_xlsplt(xlsFile, plotDir, gminy, gmaxy)
    if ivf == 1:
        boundIV.deiconify()
        boundIV.mainloop()
        iminy = float(i1.get())
        imaxy = float(i2.get())
        boundIV.withdraw()
        iv_xlsplt(xlsFile, plotDir, iminy, imaxy)
    if hpf == 1:
        messagebox.showinfo(windowTitle, "Please include only times you wish to plot. This program does not yet have user-given time bounds for this mode yet.")
        boundHP.deiconify()
        boundHP.mainloop()
        hminy = float(h1.get())
        hmaxy = float(h2.get())
        boundHP.withdraw()
        plt_Coexist_Pressure(xlsFile, plotDir, hminy, hmaxy)
    root.deiconify()
    root.mainloop()
