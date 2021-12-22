
import os 
import re
import os.path
from PyPDF2 import PdfFileMerger 
import tkinter
from tkinter import messagebox, ttk, filedialog 
from ttkthemes import ThemedTk
from pathlib import Path
from datetime import datetime
import sys

##########      TKINTER INITIATOR       ##########
root = ThemedTk(theme="aqua")
root.title('Bulk PDF Merger v1.0')
try:
    root.iconbitmap(r"icon path")
except:
    pass

##########      PROGRAM VARIABLES       ##########

loan_numbers = []
file_name_options = []

downloads_path = str(Path.home() / "Downloads")
downloads_folder= downloads_path.replace('\\\\', '\\')
todays_date = datetime.today().strftime('%m%d%y')

##########      SUB FUNCTION        ##########
def clearload():
    folderPathEntry.delete(0,'end')
    loadButton["state"] = "enable"

def browse_cmd():
    """Opens file explorer browse dialogue box for user to search for files in GUI."""
    clearload()
    root.filename = filedialog.askdirectory()
    folderPathEntry.insert(0, root.filename)
    return None


def load_folder():
    global selected_folder
    selected_folder = folderPathEntry.get()
    
    if selected_folder =='':
        tkinter.messagebox.showerror('Empty Folder Path', 'Please select the file path for the pdf folder.')
        return
    
    for file in os.listdir(selected_folder):
        if '.pdf' not in file:
            tkinter.messagebox.showerror('Invalid File Type', 'The folder you selected contains a none PDF file.\nSelect a folder that only contains the PDFs you wish to merge.')
            clearload()
            return
    
    global loan_numbers
    global file_name_options
    
    loan_numbers = list(set([re.match('\d{10}', file).group(0) for file in os.listdir(folderPathEntry.get())]))
    file_name_options = list(set([file.split(re.match('\d{10}', file).group(0))[1] for file in os.listdir(folderPathEntry.get())]))
    
    addtional_GUI()


def addtional_GUI():
    global orderLabel
    global orderOption1
    global orderOption2
    global orderOption3
    global namingEntry
    global mergeButton
    
    orderLabel = ttk.Label(root, text="Select the order of the PDFs to be merged.")
    orderLabel.grid(row=4, column=0, pady=10, padx=10)

    orderOption1_Label = ttk.Label(root, text="PDF Option 1 (required):")
    orderOption1_Label.grid(row=5, column=0, pady=10, padx=10)

    orderOption1 = ttk.Combobox(root, values=file_name_options, width=20, state="readonly")
    orderOption1.grid(row=5, column=1, pady=10, padx=10)

    orderOption2_Label = ttk.Label(root, text="PDF Option 2 (required):")
    orderOption2_Label.grid(row=6, column=0, pady=10, padx=10)

    orderOption2 = ttk.Combobox(root, values=file_name_options, width=20, state="readonly")
    orderOption2.grid(row=6, column=1, pady=10, padx=10)

    orderOption3_Label = ttk.Label(root, text="PDF Option 3 (optional):")
    orderOption3_Label.grid(row=7, column=0, pady=10, padx=10)

    orderOption3 = ttk.Combobox(root, values=file_name_options, width=20, state="readonly")
    orderOption3.grid(row=7, column=1, pady=10, padx=10)

    namingLabel = ttk.Label(root, text="Merged file naming convention. (required)\nThe loan # will be placed in front of the naming convention.")
    namingLabel.grid(row=8, column=0, pady=10, padx=10)

    namingEntry = ttk.Entry(root, width=50 )
    namingEntry.grid(row=9, column=0, pady=10, padx=10)

    mergeButton = ttk.Button(root, text='Merge', command= main_func)
    mergeButton.grid(row=10, column=0, padx=10, pady=10)


def clearall():
    folderPathEntry.delete(0,'end')
    namingEntry.delete(0,'end')
    orderOption1.set('')
    orderOption2.set('')
    orderOption3.set('')
    loadButton["state"] = "enable"
    python = sys.executable
    os.execl(python, python, * sys.argv)





##########      MAIN FUNCTION       ##########
def main_func():
    loadButton["state"] = "disable"
    mergeButton["state"] = "disable"
    option1 = orderOption1.get()
    option2 = orderOption2.get()
    option3 = orderOption3.get()
    naming = namingEntry.get()
    
    
    
    if option1 not in file_name_options or option2 not in file_name_options:
        tkinter.messagebox.showerror('Missing PDF Option', 'Please make sure an option was selected \nfor PDF option 1 and 2.')
        return
    
    if naming=="":
        tkinter.messagebox.showerror('Empty Naming Convention', 'Please provide a naming convention for the merged PDFs.')
        mergeButton["state"] = "enable"
        return

    if len(set([option1, option2, option3])) != len([option1, option2, option3]):
        tkinter.messagebox.showerror('Duplicate Selection: PDF Options', 'Please make sure you did not select the same PDF option twice.')
        mergeButton["state"] = "enable"
        return
    
    if '.pdf' not in naming:
        naming = naming + '.pdf'
    
    global selected_folder
    if '\\' in selected_folder:
        selected_folder = selected_folder + '\\'
    else:
        selected_folder = selected_folder + '/'
    
    final_folder = downloads_folder+f'\Merged_PDFs_{todays_date}'
    
    try:
        os.mkdir(final_folder)
    except:
        pass

    os.chdir(final_folder)
    
    for loan in loan_numbers:
        merger = PdfFileMerger(strict=False)
        try:
            pdf_1 = ''.join([selected_folder, loan, option1])
            pdf_2 = ''.join([selected_folder, loan, option2])
            merger.append(pdf_1)
            merger.append(pdf_2)
            if option3 in file_name_options:
                pdf_3 =''.join([selected_folder, loan, option3])
                merger.append(pdf_3)
            merger.write("{0}{1}".format(loan,naming))
            merger.close()
        except Exception as e:
            tkinter.messagebox.showerror('PDF Merge Error', f'Possible file not found.\nPleae make sure all PDF options are present for all loans.\nError message: {e}')
            loadButton["state"] = "enable"
            mergeButton["state"] = "enable"
            return
    
    clearall()


##########      MAIN GUI DESIGN      ##########
folderpathLabel = ttk.Label(root, text="Enter file path to the pdf's folder.")
folderpathLabel.grid(row=0, column=0, pady=10, padx=10)

folderPathEntry = ttk.Entry(root, width=50 )
folderPathEntry.grid(row=1, column=0, pady=10, padx=10)

browseButton = ttk.Button(root, text='Browse', command= browse_cmd)
browseButton.grid(row=1, column=1, pady=10, padx=10)

loadButton = ttk.Button(root, text='Load', command=load_folder)
loadButton.grid(row=2, column=0, padx=10, pady=10)

root.mainloop()