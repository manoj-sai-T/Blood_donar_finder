from tkinter import *
from tkinter import ttk
import tkinter
import openpyxl

window = tkinter.Tk()
window.title('Request blood')
window.configure(bg='lightblue')

app_width = 1000
app_height = 500

screen_width = window.winfo_screenwidth()
screen_height = window.winfo_screenheight()

x = (screen_width // 2) - (app_width // 2)
y = (screen_height // 2) - (app_height // 2)

window.geometry(f"{app_width}x{app_height}+{x}+{y}")


def load_data():
    bg = blood_group_combobox.get()
    path = r'C:\Users\Personal\OneDrive\Desktop\VS code\vs_python\blood_donar finder\excelfiles\bloodgrps.xlsx'
    Workbook = openpyxl.load_workbook(path)
    # sheet = Workbook.active
    if bg == 'O+':
        sheet = Workbook['O+']
    elif bg == 'O-':
        sheet = Workbook['O-']
    elif bg == 'A+':
        sheet = Workbook['A+']
    elif bg == 'A-':
        sheet = Workbook['A-']
    elif bg == 'B-':
        sheet = Workbook['B-']
    elif bg == 'B+':
        sheet = Workbook['B+']
    elif bg == 'AB-':
        sheet = Workbook['AB-']
    elif bg == 'AB+':
        sheet = Workbook['AB+']

    list_values = list(sheet.values)
    cols = list_values[0]
    tree = ttk.Treeview(window, column=cols, show="headings")
    for col_name in cols:
        tree.heading(col_name, text=col_name)
    for value_tuple in list_values[1:]:
        tree.insert('', tkinter.END, values=value_tuple)

    tree.pack(expand=True, fill='y')


label1 = Label(window, text="NAME ", font='timesnewroman,bold')
label1.pack()
entry1 = Entry(window, width=30, borderwidth=4)
entry1.pack()

label2 = Label(window, text="Blood Group", font='timesnewroman,bold')
label2.pack()

# Adding the combobox for blood group selection
blood_groups = ['A+', 'A-', 'B+', 'B-', 'O+', 'O-', 'AB+', 'AB-']
blood_group_combobox = ttk.Combobox(window, values=blood_groups, width=27)
blood_group_combobox.pack()

label3 = Label(window, text="Phone Number ", font='timesnewroman,bold')
label3.pack()
entry3 = Entry(window, width=30, borderwidth=4)
entry3.pack()

b1 = Button(text="Request", font='bold', padx=50, pady=10, width=10, bg='lightgreen', command=load_data)
b1.pack()

window.mainloop()
