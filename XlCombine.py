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

#obtian current working directory
cwd = os.getcwd()

#create root for tk
root = tk.Tk()

#tk datatypes for holding the path info
inPath = tk.StringVar()

outPath = tk.StringVar()

#setting the window title
root.title("Xl Combine")

'''
Function excelCOllate() takes no parameters and returns no parameters.
This function uses package Pandas to create a single excel spreadsheet
from many source spreadsheets
''' 
def excelCollate():
  
  #receiving the i/o paths
  ipath = inPath.get()
  
  opath = outPath.get()

  #setting the path for glob to obtain all files
  path = ipath + '*.xlsx'

  #lists to store the filepaths and the intermediate DataFrames
  FPATH = []

  frame = []

  #loop for storing the filenames
  for filename in glob.glob(path): 
    FPATH.append(filename)
    
  #loop for reading the files and adding the generated DataFrames to the list of frames
  for filename in FPATH:
    frame.append( pd.read_excel( filename ) )
    
  try:
  
    #storing the column names of one DataFrame in a list for column order
    cols = list(frame[1])
    
    #intermediate dataframe will hold data initially
    idf = pd.DataFrame()

    #loop appends all DataFrames to the intermediate DataFrame
    for f in frame:
      idf = idf.append( f )

    #store the same data in a new dataframe, but sort for column order
    df = pd.DataFrame( idf, columns=cols )

    #using a Pandas excel writer to write the finished file
    writer = pd.ExcelWriter(opath)

    df.to_excel(writer)  

    writer.save()
    
    successMessage()
    
  except IndexError:
    errorMessage()

  

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
outDText = tk.Label(root, height=1, width=10, text="Output Dir")
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


'''
Function select() runs excelCollate() when the button is pressed and if the 
directories are not empty
'''
def select():
  if (len(inPath.get()) > 0)  and (len(outPath.get()) > 0):
     excelCollate()

 
def successMessage():
  #messagebox with success message
  mb.showinfo(title='Success!', message='Successfully wrote file at ' + outPath.get())

def errorMessage():
  mb.showerror(title="Error", message="Invalid File or Directory")
 
#button to activate the collate script 
sbut = tk.Button(root, text='Submit', command=select)
sbut.grid(row=3, column=1)

root.mainloop()