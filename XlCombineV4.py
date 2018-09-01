'''
XlCombine uses Pandas to combine many excel spreadsheets into
one. This program uses a GUI to get the user's desired file locations.
'''


#required packages
import tkinter as tk
from tkinter import messagebox as mb
from tkinter import filedialog as fd
import os
import glob
import pandas as pd
from functools import partial

#obtian current working directory
cwd = os.getcwd()

#create root for tk
root = tk.Tk()

#tk datatypes for holding the path info
inPath = tk.StringVar()

outPath = tk.StringVar()

colPath = tk.StringVar()

#tk datatypes for holding the checkbox info

remDupes = tk.IntVar()

dupeNames = []

#setting the window title
root.title("Xl Combine")



'''
Function excelCollate() takes no parameters and returns no parameters.
This function uses package Pandas to create a single excel spreadsheet
from many source spreadsheets
''' 


def excelCollate( dupeNames ):
  
  #receiving the i/o paths
  ipath = inPath.get()
  
  opath = outPath.get()
  
  cPath = colPath.get()

  #setting the path for glob to obtain all files
  path = ipath + '*.xlsx'
  xlspath = ipath + '*.xls'
  csvpath = ipath + '*.csv'

  #lists to store the filepaths and the intermediate DataFrames
  FPATH = []
  
  xlsFPATH = []
  
  csvFPATH = []

  frame = []
  
  xlsframe = []
  
  csvframe = []
  
  cols = []

  #loop for storing the filenames
  for filename in glob.glob(path): 
    FPATH.append(filename)
  
  for filename in glob.glob(xlspath):
    xlsFPATH.append(filename)
    
  for filepath in glob.glob(csvpath):
    csvFPATH.append(filename)
  
  if(len(FPATH) + len(xlsFPATH) + len(csvFPATH)) > 1:
    
    #loop for reading the files and adding the generated DataFrames to the list of frames
    for filename in FPATH:
      frame.append( pd.read_excel( filename ) )
    
    for filename in xlsFPATH:
      frame.append( pd.read_excel( filename ) )
    
    for filename in csvFPATH:
      frame.append( pd.read_csv( filename ) )
    
    if cPath.endswith( '.csv' ):
      cols = list( pd.read_csv( cPath ) )
    else:
      cols = list( pd.read_excel( cPath ) )
    
    #intermediate dataframe will hold data initially
    idf = pd.DataFrame()

    #loop appends all DataFrames to the intermediate DataFrame
    for f in frame:
      idf = idf.append( f )

    if (remDupes.get() == 1):
      idf = idf.drop_duplicates(subset= dupeNames, keep='first')
    
    #store the same data in a new dataframe, but sort for column order
    df = pd.DataFrame( idf, columns=cols )

    #using a Pandas excel writer to write the finished file
    writer = pd.ExcelWriter(opath)

    df.to_excel(writer)  

    writer.save()
    
    successMessage()

  else:
    dirError()


def getDupeNames():

  global dupeNames
  
  if len(colPath.get()) > 0:
    wd = tk.Toplevel(root)
    
    cPath = colPath.get()
    cols = []
    if cPath.endswith( '.csv' ):
      cols = list( pd.read_csv( cPath ) )
    else:
      cols = list( pd.read_excel( cPath ) )
      
    dupeVars = []
    dupeChecks = []
    dupeNames = []
    
    for i in range (len(cols)):
      var = tk.IntVar()
      var.set(0)
      dupeVars.append( var )
      dupeChecks.append( tk.Checkbutton( wd, text=cols[i], variable=dupeVars[i] ) )
      
      (dupeChecks[i]).pack()
    
    def dupePrc(cols, dupeVars, dupeNames):
      for i in range (len(cols)):
        if (dupeVars[i]).get() == 1:
          dupeNames.append( cols[i] )
          
      colMessage()
      
      wd.destroy()



    dupeSubmit = partial(dupePrc, cols, dupeVars, dupeNames)
    dupeSubButton = tk.Button(wd, text='Submit', command= dupeSubmit)
    dupeSubButton.pack()
  else:
    fileError()
  
dupeButton = tk.Button(root, text="Select Duplicate Criteria Columns", command=getDupeNames )
dupeButton.grid(row=4, column=1)
  

#tk Message for the program name and main message
T = tk.Message(root, justify=tk.CENTER, text="Xl Combine\n Please Select the Data Directory\n")
T.config(font=('times', 12))
T.grid(row=0, column=1)

#tk Text for the Data Dir label
dataDText = tk.Label(root, height=1, width=10, text="Data Dir")
dataDText.grid(row=1, column=0)

#tk Text for the input Dir field
K = tk.Text(root, height=1, width = 80)
K.grid(row=1, column=1)


'''
Function getiDir() uses a "Browse" dialog to get the desired directory for
the input data
'''
def getiDir():
  #uses tk's filedialog to ask for a data directory
  inPath.set(fd.askdirectory(parent=root, initialdir=cwd, title='Please select a directory'))
  
  #checks if the path exists, if so continues
  if len(inPath.get()) > 0:
    if not(inPath.get().endswith('/')):
      inPath.set(inPath.get() + '/')
    
    K.delete(1.0, tk.END)
    K.insert(tk.END, inPath.get())

#button to activate the browse dialog
ibut = tk.Button(root, text="Browse", command=getiDir)
ibut.grid(row=1, column=2)

#tk text for the Out Dir label
outDText = tk.Label(root, height=1, width=10, text="Output File")
outDText.grid(row=2, column=0)

#tk text for the Out Dir field
M = tk.Text(root, height=1, width = 80)
M.grid(row=2, column=1)

'''
Function getiDir() uses a "Browse" dialog to get the desired directory for
the output file
'''
def getoDir():
  #uses tk's filedialog to ask for the save as directory
  outPath.set(fd.asksaveasfilename(parent=root, initialdir=cwd, title='Save As', filetypes = (("Microsoft Excel", "*.xlsx"), ("all files", "*.*"))))
  
  #checks if the path exists, if so continues
  if len(outPath.get()) > 0:
    if not(outPath.get().endswith('.xlsx')):
      outPath.set(outPath.get() + '.xlsx')
    
    M.delete(1.0, tk.END)
    M.insert(tk.END, outPath.get())


#button to activate the saveas dialog
obut = tk.Button(root, text="Browse", command=getoDir)
obut.grid(row=2, column=2)

colDText = tk.Label(root, height=1, width=10, text="Column File")
colDText.grid(row=3, column=0)

N = tk.Text(root, height=1, width=80)
N.grid(row=3, column=1)

def getcDir():
  
  colPath.set(fd.askopenfilename(parent=root, initialdir=cwd, title="Select File for Column Names", filetypes = (("Microsoft Excel", "*.xlsx"),("Excel 97-03", "*.xls"), ("Comma Separated Data", "*.csv"), ("all files", "*.*"))))
  
  if len(colPath.get()) > 0:
    N.delete(1.0, tk.END)
    N.insert(tk.END, colPath.get())
    
cbut = tk.Button(root, text="Browse", command = getcDir)
cbut.grid(row=3, column=2)

'''
Function select() runs excelCollate() when the button is pressed and if the 
directories are not empty
'''
def select():
 
  if ( (len(inPath.get()) > 0)  and (len(outPath.get()) > 0) ) and len(colPath.get()) > 0:
     excelCollate( dupeNames )

  else:
    errorMessage()

 
def successMessage():
  #messagebox with success message
  mb.showinfo(title='Success!', message='Successfully wrote file at ' + outPath.get())
  
def colMessage():
  #messagebox with success message
  mb.showinfo(title='Success!', message='Column Names Saved')

def errorMessage():
  mb.showerror(title="Error", message="Please Choose Directories And Files")
  
def dirError():
  mb.showerror(title="Error", message="Data Directory is Empty or Only has One File")
  
def fileError():
  mb.showerror(title="Error", message="Invalid or Missing Column File")
 
#button to activate the collate script 
sbut = tk.Button(root, text='Submit', command=select)
sbut.grid(row=5, column=1)


#checkbox for choosing whether to remove duplicates
dCheck = tk.Checkbutton(root, text="Remove Duplicates?", variable = remDupes)
dCheck.grid(row=4, column=0)


root.mainloop()