"""
    CopyRight Dr. Ahmad Hamdi Emara 2020
"""
import tkinter as tk
from tkinter import filedialog
# from tkinter import messagebox
# import sys
import pandas as pd
import numpy as np

root = tk.Tk()

def main():
    main_canvas = tk.Canvas(root, width = 300, height = 300, bg = 'lightsteelblue2', relief = 'raised')
    main_canvas.pack()

    
    root.title('Google Contacts Conversion Tool')
    center(root)

    lbl = tk.Label(root, text='Google Contacts \nConversion Tool', bg = 'lightsteelblue2', fg='blue')
    lbl.config(font=('helvetica', 20))
    main_canvas.create_window(150, 60, window=lbl)

    browseButton_Excel = tk.Button(text="      Import excel file     ", command=getExcel, bg='green', fg='white', font=('helvetica', 12, 'bold'))
    main_canvas.create_window(150, 130, window=browseButton_Excel)

    saveAsButton_CSV = tk.Button(text='Convert to contacts', command=convertToCSV, bg='green', fg='white', font=('helvetica', 13, 'bold'))
    main_canvas.create_window(150, 180, window=saveAsButton_CSV)

    closeAppButton = tk.Button(text='Exit app', command=closeApp, bg='red', fg='white', font=('helvetica', 14, 'bold'))
    main_canvas.create_window(150, 240, window=closeAppButton)

    lbl = tk.Label(root, text='By Dr. Ahmad Hamdi Emara', bg = 'lightsteelblue3', fg='blue')
    lbl.config(font=('helvetica', 12))
    main_canvas.create_window(150, 280, window=lbl)

    root.mainloop()


def getExcel ():
    global read_file
    
    import_file_path = filedialog.askopenfilename()
    read_file = pd.read_excel(import_file_path, dtype={'MOBILE NO.':str})
    


def convertToCSV ():
    global read_file
    
    export_file_path = filedialog.asksaveasfilename(defaultextension='.csv', title = "Save your contacts", initialfile = "contacts")
    
    # edit read file here before saving it to csv
    # change column names to fit google contacts format
    read_file.columns = ["Id", "Gender", "Name", "Middle name", "Family name", "Phone 1 - Value", "Receipts", "Amount", "City", "Area", "Street", "Branch", "Language", "Delivery", "Active", "Credit"]

    # convert any entry with empty phone number to np.nan for later dropping.
    read_file['Phone 1 - Value'].replace('', np.nan, inplace=True)

    # drop unnecessary columns from the data frame.
    read_file = read_file.drop(["Receipts", "Amount", "Delivery", "Active", "Credit", "Id"], axis=1)
    # drop completely empty entries.
    read_file.dropna(subset=['Phone 1 - Value'], inplace=True)
    # drop duplicates
    read_file.drop_duplicates()
    
    # drop last row
    read_file.drop(read_file.tail(1).index, inplace=True) 

    # add the middle name to the first name.
    read_file['Name'] = read_file['Name'] + ' ' + read_file['Middle name']

    # insert the "phone type" column and assign every value to "Mobile" before each phone number in the data frame.
    read_file.insert(4, 'Phone 1 - Type', 'Mobile', allow_duplicates = True)
    print(read_file.head())

    # save the ready google contacts file.
    read_file.to_csv(export_file_path, index = None, header=True, index_label = True, encoding='utf-8')
    # root.destroy()


def center(win):
    """
    centers a tkinter window
    :param win: the root or Toplevel window to center
    """
    win.update_idletasks()
    width = win.winfo_width()
    frm_width = win.winfo_rootx() - win.winfo_x()
    win_width = width + 2 * frm_width
    height = win.winfo_height()
    titlebar_height = win.winfo_rooty() - win.winfo_y()
    win_height = height + titlebar_height + frm_width
    x = win.winfo_screenwidth() // 2 - win_width // 2
    y = win.winfo_screenheight() // 2 - win_height // 2
    win.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    win.deiconify()

def closeApp():
    root.destroy()
    
if __name__ == "__main__":
    main()
