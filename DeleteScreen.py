from tkinter import *
import openpyxl

root = Tk()
delete_list = {}
def deleteRoute():

    def confirmDelete():
        global delete_list
        for key, value in delete_list.items():
            if value:
                boulder_sheet.delete_rows(key)
                delete_list = [row - 1 for row in delete_list]
        route_book.save('Routes.xlsx')

    def btnPress(num):
        if button_list[num-1].cget('fg') == 'red':
            delete_list[num] = False
            button_list[num-1].configure(fg='green')
        else:
            delete_list[num] = True
            button_list[num-1].configure(fg='red')

    route_book = openpyxl.load_workbook('Routes.xlsx')
    boulder_sheet = route_book['Boulders']

    delete_window = Tk()

    button_list = []


    num_buttons = 0

    info_label = Label(delete_window, text='Pressing confirm will delete all routes highlighted red')
    info_label.grid(row=0, column=1)

    for cell in range(2, boulder_sheet.max_row + 1):
        grade = boulder_sheet.cell(row=cell, column=1).value
        wall = boulder_sheet.cell(row=cell, column=2).value
        setter = boulder_sheet.cell(row=cell, column=3).value
        color = boulder_sheet.cell(row=cell, column=4).value

        if setter == setterDrop.get():
            num_buttons += 1
            button_list.append(Button(delete_window, fg='red', text='{}, {}, {}, {}'.format(color, grade, wall, setter),
                                     command=lambda x=num_buttons: btnPress(x)))
            button_list[num_buttons - 1].grid(row=num_buttons, column=1)
            delete_list.update({num_buttons: True})

    confirm_button = Button(delete_window, text='Confirm Deletion', command=confirmDelete)
    confirm_button.grid(row=num_buttons+1, column=1)

wall_options = ['Wall:', 'Topout', 'Main']
setter_options = ['Setter:', 'Chase', 'Chris', 'Christian', 'Jeff', 'Jeremy', 'Joey', 'Mitch']
color_options = ['Color:', 'Green', 'Orange', 'Pink']

wallDrop = StringVar()
wallDrop.set(wall_options[0])

setterDrop = StringVar()
setterDrop.set(setter_options[0])

colorDrop = StringVar()
colorDrop.set(color_options[0])


wallMenu = OptionMenu(root, wallDrop, *wall_options)
wallMenu.grid(row=0, column=1)

setterMenu = OptionMenu(root, setterDrop, *setter_options)
setterMenu.grid(row=0, column=2)

colorMenu = OptionMenu(root, colorDrop, *color_options)
colorMenu.grid(row=0, column=3)

submitButton = Button(root, text='Search For Routes', command=deleteRoute)
submitButton.grid(row=1, column=1, columnspan=2)

root.mainloop()
