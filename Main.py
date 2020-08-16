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

# Delete window
def deleteWindow():
    del_root = tkinter.Toplevel(root)
    del_root.geometry('300x400')

    global wall_options
    del_wall_options = wall_options

    def deleteRoute():
        global delete_list

        def confirmDelete():
            global delete_list
            deleted = 0
            for key, value in delete_list.items():
                if value:
                    deleted += 1

            # Confirm deletion
            confirm_answer = tkinter.messagebox.askquestion('Confirm Deletion', 'Your are about to delete {} '
                                                                                'Routes. Continue?'.format(deleted))

            if confirm_answer == 'yes':
                for key2, value2 in delete_list.items():
                    if value2:
                        boulder_sheet.delete_rows(key2)
                        delete_list = {key2 + 1: value2}
                tkinter.messagebox.showinfo('', 'Deleted {} routes successfully.'.format(deleted))
                delete_window.destroy()
            else:
                tkinter.messagebox.showinfo('Successful Cancel', 'Canceled route deletion.')
            route_book.save('Routes.xlsx')

        # Change buttons to represent excluded routes
        def btnPress(num, row):
            if button_list[num - 1].cget('fg') == 'red':
                delete_list[row] = False
                button_list[num - 1].configure(fg='green')
            else:
                delete_list[row] = True
                button_list[num - 1].configure(fg='red')

        route_book = openpyxl.load_workbook('Routes.xlsx')
        boulder_sheet = route_book['Boulders']

        delete_window = tkinter.Toplevel(root)

        button_list = []

        num_buttons = 0

        info_label = Label(delete_window, text='Pressing confirm will delete all routes highlighted red')
        info_label.grid(row=0, column=1)

        for cell in range(1, boulder_sheet.max_row + 1):
            grade = boulder_sheet.cell(row=cell, column=1).value
            wall = boulder_sheet.cell(row=cell, column=2).value
            setter = boulder_sheet.cell(row=cell, column=3).value
            color = boulder_sheet.cell(row=cell, column=4).value

            if wall == delWallDrop.get():
                num_buttons += 1
                button_list.append(
                    Button(delete_window, fg='red', text='{}, {}, {}, {}'.format(color, grade, wall, setter),
                           command=lambda x=num_buttons, y=cell: btnPress(x, y)))
                button_list[num_buttons - 1].grid(row=num_buttons, column=1)
                delete_list.update({cell: True})

        del_confirm_button = Button(delete_window, text='Confirm Deletion', command=confirmDelete)
        del_confirm_button.grid(row=num_buttons + 1, column=1)

    delWallDrop = StringVar()
    delWallDrop.set(del_wall_options[0])

    wallDelMenu = OptionMenu(del_root, delWallDrop, *del_wall_options)
    wallDelMenu.grid(row=0, column=1)

    delSubmitButton = Button(del_root, text='Search For Routes', command=deleteRoute)
    delSubmitButton.grid(row=1, column=1, columnspan=2)

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

        g_graph_root.destroy()

        def setRopeGraph():
            # Update spreadsheet
            data_sheet.cell(row=2, column=2).value = int(entry_6.get())
            data_sheet.cell(row=3, column=2).value = int(entry_7.get())
            data_sheet.cell(row=4, column=2).value = int(entry_8.get())
            data_sheet.cell(row=5, column=2).value = int(entry_9.get())
            data_sheet.cell(row=6, column=2).value = int(entry_10.get())
            data_sheet.cell(row=7, column=2).value = int(entry_11.get())
            data_sheet.cell(row=8, column=2).value = int(entry_12.get())
            data_sheet.cell(row=9, column=2).value = int(entry_13.get())

            route_book.save('Routes.xlsx')
            rope_graph_root.destroy()

        # Creates a window with entry boxes for every grade type
        rope_graph_root = tkinter.Toplevel(root)

        main_label = Label(rope_graph_root, text='Enter your preferred route curve below:')
        main_label.grid(row=0, column=0, columnspan=2)

        label_6 = Label(rope_graph_root, text='5.6s')
        label_6.grid(row=1, column=0)
        entry_6 = Entry(rope_graph_root)
        entry_6.grid(row=1, column=1)

        label_7 = Label(rope_graph_root, text='5.7s')
        label_7.grid(row=2, column=0)
        entry_7 = Entry(rope_graph_root)
        entry_7.grid(row=2, column=1)

        label_8 = Label(rope_graph_root, text='5.8s')
        label_8.grid(row=3, column=0)
        entry_8 = Entry(rope_graph_root)
        entry_8.grid(row=3, column=1)

        label_9 = Label(rope_graph_root, text='5.9s')
        label_9.grid(row=4, column=0)
        entry_9 = Entry(rope_graph_root)
        entry_9.grid(row=4, column=1)

        label_10 = Label(rope_graph_root, text='5.10s')
        label_10.grid(row=5, column=0)
        entry_10 = Entry(rope_graph_root)
        entry_10.grid(row=5, column=1)

        label_11 = Label(rope_graph_root, text='5.11s')
        label_11.grid(row=6, column=0)
        entry_11 = Entry(rope_graph_root)
        entry_11.grid(row=6, column=1)

        label_12 = Label(rope_graph_root, text='5.12s')
        label_12.grid(row=7, column=0)
        entry_12 = Entry(rope_graph_root)
        entry_12.grid(row=7, column=1)

        label_13 = Label(rope_graph_root, text='5.13s')
        label_13.grid(row=8, column=0)
        entry_13 = Entry(rope_graph_root)
        entry_13.grid(row=8, column=1)

        get_graph = Button(rope_graph_root, text='Set Graph', command=setRopeGraph)
        get_graph.grid(row=11, column=0, columnspan=2)

        rope_graph_root.mainloop()

    # Creates a window with entry boxes for every grade type
    g_graph_root = tkinter.Toplevel(root)

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
    column = data_sheet['A']
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
wall_options = ['Wall:']
setter_options = ['Setter:']
color_options = ['Color:']
delete_list = {}

route_book = openpyxl.load_workbook('Routes.xlsx')
data_sheet = route_book['Data']

for cell in range(2, data_sheet.max_row + 1):
    # Make variable out of cells
    setter_cell = data_sheet.cell(row=cell, column=3).value
    color_cell = data_sheet.cell(row=cell, column=4).value
    wall_cell = data_sheet.cell(row=cell, column=5).value

    # Iterate over variables, appending to appropriate lists
    if setter_cell is not None:
        setter_options.append(setter_cell)
    if color_cell is not None:
        color_options.append(color_cell)
    if wall_cell is not None:
        wall_options.append(wall_cell)


gradeDrop = StringVar()
gradeDrop.set(grade_options[0])

wallDrop = StringVar()
wallDrop.set(wall_options[0])

setterDrop = StringVar()
setterDrop.set(setter_options[0])

colorDrop = StringVar()
colorDrop.set(color_options[0])

gradeMenu = OptionMenu(root, gradeDrop, *grade_options)
gradeMenu.grid(row=0, column=0, sticky=NSEW)

wallMenu = OptionMenu(root, wallDrop, *wall_options)
wallMenu.grid(row=0, column=1, sticky=NSEW)

setterMenu = OptionMenu(root, setterDrop, *setter_options)
setterMenu.grid(row=0, column=2, sticky=NSEW)

colorMenu = OptionMenu(root, colorDrop, *color_options)
colorMenu.grid(row=0, column=3, sticky=NSEW)

submitButton = Button(root, text='Submit Route', command=submitRoute)
submitButton.grid(row=1, column=1, sticky=NSEW)

deleteButton = Button(root, text='Delete Routes', command=deleteWindow)
deleteButton.grid(row=1, column=2, sticky=NSEW)

goalButton = Button(root, text='Set Ideal Route Curve', command=setGoal)
goalButton.grid(row=2, column=1, columnspan=2, sticky=NSEW)

graphButton = Button(root, text='Route Graphs', command=graphRoutes)
graphButton.grid(row=3, column=1, columnspan=2, sticky=NSEW)

root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=1)
root.grid_columnconfigure(2, weight=1)
root.grid_columnconfigure(3, weight=1)

root.grid_rowconfigure(0, weight=1)
root.grid_rowconfigure(1, weight=1)
root.grid_rowconfigure(2, weight=1)
root.grid_rowconfigure(3, weight=1)
root.grid_rowconfigure(4, weight=1)

root.mainloop()
