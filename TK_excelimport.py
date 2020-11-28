# -*- coding: utf-8 -*-

from tkinter import *
from tkinter import filedialog,messagebox,ttk
import pandas as pd

window = Tk() 
window.configure(background='white')
#window.geometry('600x400')
window.title('Length Unit Converter')

e_l=Label(window, text = 'Value length: ',bg='white',anchor='w')
e_l.grid(row=2, column=0)

value_int=StringVar()
e_v=Entry(window,width=20,textvariable=value_int)
e_v.grid(row=2,column=1,padx=10,pady=10)

e_l=Label(window, text = 'Convert From: ',bg='white',anchor='w')
e_l.grid(row=3, column=0)

e_l=Label(window, text = 'Convert To: ',bg='white',anchor='w')
e_l.grid(row=4, column=0)

label_file = Label(window, text='No File Selected',bg='white',anchor='w')
label_file.grid(row=0, column=1, columnspan=4, padx=20, pady=20)

label_load = Label(window, text='No File Loaded',bg='white',anchor='w')
label_load.grid(row=1, column=1, columnspan=4, padx=20, pady=20)

def upload_excel():
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetype=(("xlsx files", "*.xlsx"),("All Files", "*.*")))
    label_file["text"] = filename
    return None


def matrix_iter(unit_from,unit_to):
    uf=None
    ut=None
    for i in range(len(matrix.index)):
        if unit_from==matrix.index[i]:
            uf=i
    for j in range(len(matrix.columns)):
        if unit_to==matrix.columns[j]:
            ut=j
    conv_fact=matrix.iat[uf,ut]
    return conv_fact


def converter():
    unit_from=f.get()
    unit_to=str(t.get())
    value_unit=float(value_int.get())
    conv_fact=matrix_iter(unit_from,unit_to)
    result=value_unit*conv_fact
    result_w.delete('1.0', END)
    result_w.insert(END,float(result))

           
def reset():
    e_v.delete(0,END)
    result_w.delete('1.0',END)
    from_v.set('')
    to_v.set('')


def load_excel_data():
    global matrix
    file_path = label_file["text"]
    try:
        excel_filename = r"{}".format(file_path)
        if excel_filename[-4:] == ".csv":
            matrix = pd.read_csv(excel_filename,index_col=0)
        else:
            matrix = pd.read_excel(excel_filename,index_col=0)
        from_v['values']=[i for i in matrix.index]
        to_v['values']=[i for i in matrix.columns]
        label_load['text']='Loaded'
    except ValueError:
        tk.messagebox.showerror("Info", "Invalid file")
        return None


f=StringVar()
from_v=ttk.Combobox(window,width=20,textvariable=f)
from_v.grid(column=1, row=3,pady=10)

t=StringVar()
to_v=ttk.Combobox(window,width=20,textvariable=t)
to_v.grid(column=1, row=4,pady=10)

button1 = ttk.Button(window, text="Browse File", command=lambda: upload_excel())
button1.grid(row=0,column=0,pady=10)

button2 = ttk.Button(window, text="Load File", command=lambda: load_excel_data())
button2.grid(row=1,column=0,pady=10)


button_convert=Button(window,text='Convert', command=converter)
button_convert.grid(column=0,row=6,pady=10)

button_reset=Button(window,text='Reset',command=reset)
button_reset.grid(column=1,row=6,pady=10)

e_l=Label(window, text = 'Result: ',bg='white',anchor='w')
e_l.grid(row=5, column=0,pady=10)

result_w=Text(window, height = 1, width = 20)
result_w.grid(row=5,column=1,pady=10)
result_w.configure(font="TkDefaultFont")

window.mainloop() 