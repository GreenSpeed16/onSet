# Imports
from tkinter import *
import tkinter.messagebox
import matplotlib.pyplot as plt
import openpyxl
import numpy as np

# Define functions
# Submit route

def submitRoute():
    # Load workbook
    route_book = openpyxl.load_workbook('Routes.xlsx')
    boulder_sheet = route_book['Boulders']

    # If all values have been selected, update spreadsheet
    if gradeDrop.get() != 'Grade:' and wallDrop.get() != 'Wall:'\
            and setterDrop.get() != 'Setter:' and colorDrop.get() != 'Color:':
        boulder_sheet.cell(column=1, row=boulder_sheet.max_row + 1).value = gradeDrop.get()
        boulder_sheet.cell(column=2, row=boulder_sheet.max_row).value = wallDrop.get()
        boulder_sheet.cell(column=3, row=boulder_sheet.max_row).value = setterDrop.get()
        boulder_sheet.cell(column=4, row=boulder_sheet.max_row).value = colorDrop.get()
        tkinter.messagebox.showinfo(title='Route Added', message='Your route has been successfully added.')
        gradeDrop.set(grade_options[0])
        wallDrop.set(wall_options[0])
        setterDrop.set(setter_options[0])
        colorDrop.set(color_options[0])
    else:
        tkinter.messagebox.showerror(title=None, message='One or more options was not filled out, please try again')

    # Save workbook
    route_book.save('Routes.xlsx')

# Goal Graph

def setGoal():
    def setGraph():
        # Open excel document
        route_book = openpyxl.load_workbook('Routes.xlsx')
        data_sheet = route_book['Data']

        # Update spreadsheet
        data_sheet.cell(row=2, column=1).value = int(entry_v0.get())
        data_sheet.cell(row=3, column=1).value = int(entry_v1.get())
        data_sheet.cell(row=4, column=1).value = int(entry_v2.get())
        data_sheet.cell(row=5, column=1).value = int(entry_v3.get())
        data_sheet.cell(row=6, column=1).value = int(entry_v4.get())
        data_sheet.cell(row=7, column=1).value = int(entry_v5.get())
        data_sheet.cell(row=8, column=1).value = int(entry_v6.get())
        data_sheet.cell(row=9, column=1).value = int(entry_v7.get())
        data_sheet.cell(row=10, column=1).value = int(entry_v8.get())
        data_sheet.cell(row=11, column=1).value = int(entry_v9.get())

        route_book.save('Routes.xlsx')
        g_graph_root.destroy()

    # Creates a window with entry boxes for every grade type
    g_graph_root = Tk()

    main_label = Label(g_graph_root, text='Enter your preferred route curve below:')
    main_label.grid(row=0, column=0, columnspan=2)

    label_v0 = Label(g_graph_root, text='V0s')
    label_v0.grid(row=1, column=0)
    entry_v0 = Entry(g_graph_root)
    entry_v0.grid(row=1, column=1)

    label_v1 = Label(g_graph_root, text='V1s')
    label_v1.grid(row=2, column=0)
    entry_v1 = Entry(g_graph_root)
    entry_v1.grid(row=2, column=1)

    label_v2 = Label(g_graph_root, text='V2s')
    label_v2.grid(row=3, column=0)
    entry_v2 = Entry(g_graph_root)
    entry_v2.grid(row=3, column=1)

    label_v3 = Label(g_graph_root, text='V3s')
    label_v3.grid(row=4, column=0)
    entry_v3 = Entry(g_graph_root)
    entry_v3.grid(row=4, column=1)

    label_v4 = Label(g_graph_root, text='V4s')
    label_v4.grid(row=5, column=0)
    entry_v4 = Entry(g_graph_root)
    entry_v4.grid(row=5, column=1)

    label_v5 = Label(g_graph_root, text='V5s')
    label_v5.grid(row=6, column=0)
    entry_v5 = Entry(g_graph_root)
    entry_v5.grid(row=6, column=1)

    label_v6 = Label(g_graph_root, text='V6s')
    label_v6.grid(row=7, column=0)
    entry_v6 = Entry(g_graph_root)
    entry_v6.grid(row=7, column=1)

    label_v7 = Label(g_graph_root, text='V7s')
    label_v7.grid(row=8, column=0)
    entry_v7 = Entry(g_graph_root)
    entry_v7.grid(row=8, column=1)

    label_v8 = Label(g_graph_root, text='V8s')
    label_v8.grid(row=9, column=0)
    entry_v8 = Entry(g_graph_root)
    entry_v8.grid(row=9, column=1)

    label_v9 = Label(g_graph_root, text='V9s')
    label_v9.grid(row=10, column=0)
    entry_v9 = Entry(g_graph_root)
    entry_v9.grid(row=10, column=1)

    get_graph = Button(g_graph_root, text='Set Graph', command=setGraph)
    get_graph.grid(row=11, column=0, columnspan=2)

    g_graph_root.mainloop()

# Show current and goal graph
def graphRoutes():
    # Get data for current graph
    route_book = openpyxl.load_workbook('Routes.xlsx')
    boulder_sheet = route_book['Boulders']

    grade_list = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    x_axis = ['V0', 'V1', 'V2', 'V3', 'V4', 'V5', 'V6', 'V7', 'V8', 'V9']
    positions = np.arange(len(grade_list))

    for cell in range(2, boulder_sheet.max_row + 1):
        temp_grade = boulder_sheet.cell(column=1, row=cell).value
        grade_int = int(temp_grade.replace('V', ''))
        # Fill the list with appropriate values
        grade_list[grade_int] += 1

    # Get data for goal graph
    goal_list = []
    data_sheet = route_book['Data']
    for cell in range(2, data_sheet.max_row + 1):
        temp_g_grade = data_sheet.cell(row=cell, column=1).value
        goal_list.append(temp_g_grade)

    route_book.save('Routes.xlsx')

    plt.bar(positions, grade_list, width=0.5, color='red')
    plt.bar(positions + 0.5, goal_list, width=0.5)
    plt.legend(['Current Spread', 'Goal Spread'])
    plt.title('Route Comparison')
    plt.xticks(positions + 0.25, x_axis)
    plt.ylabel('Amount')

    plt.show()

# Initialize tkinter module
root = Tk()
root.title('onSet')
# root.geometry('400x400')

# Define route variables and initial values
grade_options = ['Grade:', 'V0', 'V1', 'V2', 'V3', 'V4', 'V5', 'V6', 'V7', 'V8', 'V9']
wall_options = ['Wall:', 'Topout', 'Main']
setter_options = ['Setter:', 'Chase', 'Chris', 'Christian', 'Jeff', 'Jeremy', 'Joey', 'Mitch']
color_options = ['Color:', 'Green', 'Orange', 'Pink']

gradeDrop = StringVar()
gradeDrop.set(grade_options[0])

wallDrop = StringVar()
wallDrop.set(wall_options[0])

setterDrop = StringVar()
setterDrop.set(setter_options[0])

colorDrop = StringVar()
colorDrop.set(color_options[0])

gradeMenu = OptionMenu(root, gradeDrop, *grade_options)
gradeMenu.grid(row=0, column=0)

wallMenu = OptionMenu(root, wallDrop, *wall_options)
wallMenu.grid(row=0, column=1)

setterMenu = OptionMenu(root, setterDrop, *setter_options)
setterMenu.grid(row=0, column=2)

colorMenu = OptionMenu(root, colorDrop, *color_options)
colorMenu.grid(row=0, column=3)

submitButton = Button(root, text='Submit Route', command=submitRoute)
submitButton.grid(row=1, column=1, columnspan=2)

goalButton = Button(root, text='Set Ideal Route Curve', command=setGoal)
goalButton.grid(row=3, column=1, columnspan=2)

graphButton = Button(root, text='Route Graphs', command=graphRoutes)
graphButton.grid(row=5, column=1, columnspan=2)

root.mainloop()
