import tkinter as tk
from tkinter import filedialog as fd
import pandas as pd
import re

root = tk.Tk()
root.title('CSV Converter')
root.resizable(False, False)
root.geometry('600x400')

labelResultado = tk.Label(root, font=("Arial", 12), text='Select the CSV file:')
labelResultado.pack(pady=10)

def selectFile():
  global fileCsv
  fileCsv = fd.askopenfilename(
        title='Select the CSV file',
        initialdir='',
        filetypes=[("CSV Files", "*.csv")])
  fileName = tk.Label(root, font=("Arial", 12), text=fileCsv)
  fileName.pack(pady=10)
  
selectBtn = tk.Button(root, text='Select', command=selectFile)
selectBtn.pack(pady=10)

def converter():
  df = pd.read_csv(fileCsv)

  lista=[]

  for i in range(len(df)):
    lista.append([i+1, df['Name'][i], re.sub(r"\s*\[.*?\]\s*", "", df["Platform"][i])])

  df_export = pd.DataFrame(lista, columns=['Index', 'Game', 'Platform'])

  with pd.ExcelWriter('VGCollection.xlsx', engine='xlsxwriter') as writer:
      df_export.to_excel(writer, sheet_name='Collection', index=False)

      workbook  = writer.book
      worksheet = writer.sheets['Collection']

      worksheet.set_column('A:A', 5)
      worksheet.set_column('B:B', 60) 
      worksheet.set_column('C:C', 21)
  
  sucess = tk.Label(root, font=("Arial", 12), text='File converted', fg="green")
  sucess.pack(pady=10)
    
converterBtn = tk.Button(root, text='Convert', command=converter)
converterBtn.pack(pady=10)

root.mainloop()


