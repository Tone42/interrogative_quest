# Database fields entry program

import tkinter as tk

root=tk.Tk()

# setting windows size
root.geometry("600x400")

# declaring string variable
# for entering information into the database

who_field = tk.StringVar()
what_field = tk.StringVar()
when_field = tk.StringVar()
where_field = tk.StringVar()
why_field = tk.StringVar()
how_field = tk.StringVar()

# defining the function that will
# get the what and why etc. and
# print them on the screen
def submit():

    who = who_field.get()
    what = what_field.get()
    when = when_field.get()
    where = where_field.get()
    why = why_field.get()
    how = how_field.get()

    #Code for excel

    from fileinput import filename
    from openpyxl import Workbook

    #Creates file
    filename = "interogative_quest1.xlsx"

    #Allows to work with fields
    workbook = Workbook()
    sheet = workbook.active

    sheet["A1"] = who
    sheet["B1"] = what
    sheet["C1"] = when
    sheet["D1"] = where
    sheet["E1"] = why
    sheet["F1"] = how

    #save duh
    workbook.save(filename=filename)
	
    #This is what sets the fields when submit is pressed
    who_field.set("")
    what_field.set("")
    when_field.set("")
    where_field.set("")
    why_field.set("")
    how_field.set("")
	
# creating a label using widget Label
# and
# creating a entry using widget Entry

who_label = tk.Label(root, text = 'Who' , font=('calibre',10, 'bold'))
who_entry = tk.Entry(root,textvariable = who_field, font=('calibre',10,'normal'))

what_label = tk.Label(root, text = 'What', font=('calibre',10, 'bold'))
what_entry = tk.Entry(root,textvariable = what_field, font=('calibre',10,'normal'))

when_label = tk.Label(root, text = 'When', font=('calibre',10, 'bold'))
when_entry = tk.Entry(root,textvariable = when_field, font=('calibre',10,'normal'))


where_label = tk.Label(root, text = 'Where', font=('calibre',10, 'bold'))
where_entry = tk.Entry(root,textvariable = where_field, font=('calibre',10,'normal'))

why_label = tk.Label(root, text = 'Why', font = ('calibre',10,'bold'))
why_entry=tk.Entry(root, textvariable = why_field, font = ('calibre',10,'normal'))

how_label = tk.Label(root, text = 'How', font=('calibre',10, 'bold'))
how_entry=tk.Entry(root, textvariable = how_field, font = ('calibre',10,'normal'))



# creates a button using the widgetthat will call the submit function

sub_btn=tk.Button(root,text = 'Submit', command = submit)

# placing the label and entry in
# the required position using grid
# method

who_label.grid(row=0, column=0)
who_entry.grid(row=0, column=1)

what_label.grid(row=1,column=0)
what_entry.grid(row=1,column=1)

when_label.grid(row=2, column=0)
when_entry.grid(row=2, column=1)

where_label.grid(row=3, column=0)
where_entry.grid(row=3, column=1)

why_label.grid(row=4,column=0)
why_entry.grid(row=4,column=1)

how_label.grid(row=5, column=0)
how_entry.grid(row=5, column=1)

sub_btn.grid(row=7,column=1)

#infinite loop
root.mainloop()



