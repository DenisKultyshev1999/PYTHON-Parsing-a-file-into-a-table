from tkinter import *
import pandas as pd
from tkinter import ttk, filedialog
import openpyxl
from jsonpath_rw import jsonpath, parse
from collections import Counter
import json
import xlwt




root = Tk()
root.title("Table") 
root.geometry("800x500")

data=json.load(open('data.json'))


list_creation_date = []                     #creating empty lists
list_title = []
list_display_name = []
list_is_answered = []
list_link = []

i=0
for item in data['items']:                         #filling lists with data from a file
    list_creation_date .append(item['creation_date'])
    list_title.append(item['title'])
    list_display_name.append(data['items'][i]['owner']['display_name'])
    i+=1
    list_is_answered.append(item['is_answered'])
    list_link.append(item['link'])

#creating an xsl file and filling it with data from arrays for subsequent visualization
book = xlwt.Workbook()
sheet = book.add_sheet("Sheet")

cols = ["A", "B", "C", "D", "E"]
sheet.row(0).write(0, "Creation date")
sheet.row(0).write(1, "Title")
sheet.row(0).write(2, "Author")
sheet.row(0).write(3, "Answered?")
sheet.row(0).write(4, "Link")

for j in range(1,i):
    sheet.row(j).write(0,list_creation_date[j-1])
    sheet.row(j).write(1,list_title[j-1])
    sheet.row(j).write(2,list_display_name[j-1])
    sheet.row(j).write(3,list_is_answered[j-1])
    sheet.row(j).write(4,list_link[j-1])

book.save("test.xls")

my_frame = Frame(root)   #Create frame
my_frame.pack(pady=20)

my_tree = ttk.Treeview()  #Create treeview

def file_open():     #File open function

     title = "Open A File"

     df = pd.read_excel('test.xls')
     clear_tree()  #Clear old treeview

     my_tree["column"] = list(df.columns) # Set up new treeview
     my_tree["show"] = "headings"

     for column in my_tree[ "column"]:       #Loop thru column list for headers
         my_tree.heading(column, text=column)

     df_rows = df.to_numpy().tolist()    #Put data in treeview
     for row in df_rows:
         my_tree.insert("", "end", values=row)

     my_tree.pack() #Pack the treeview finally


def clear_tree():
    my_tree.delete(*my_tree.get_children())



my_menu = Menu(root)  #Add a menu
root.config(menu=my_menu)

file_menu = Menu(my_menu, tearoff=False)             #Add menu dropdown
my_menu.add_cascade(label="Spreadsheets", menu=file_menu)
file_menu.add_command(label="Open", command=file_open)

my_label = Label(root, text='')
my_label.pack(pady=20)

root.mainloop()
