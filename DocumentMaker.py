from docxtpl import DocxTemplate
from tkinter import *
root = Tk()
root.title('DocumentChanger')
root.minsize(800, 400)
root.configure(background='#DBD7D6')
upperText = Label(fg='black', bg ='grey', width=50, height=1, text = 'Введите название документа')
enterFilename = Entry(width = 50)
count = 0
enteredVarName = []
enteredVarValue = []

change1Var = Entry(width = 40)
chenge1Value = Entry(width = 40)
Label(fg='black', bg='grey', width=30, height=1, text='Введите имя переменной ').grid(column=0, row = 4)
Label(fg='black', bg='grey', width=30, height=1, text='Введите текст переменной ').grid(column=1, row = 4)


def DocChenger(event):
    filename = str(enterFilename.get())
    doc = DocxTemplate(filename)
    context = {}
    for i in range(0, len(enteredVarName)):
        context[str(enteredVarName[i].get())] = str(enteredVarValue[i].get())

    doc.render(context)
    doc.save('Generated-' + filename)

def add_line(event):
    global count
    enteryName = Entry()
    enteredVarName.append(enteryName)
    enteryName.grid(row = 5 + count)
    enteryValue = Entry()
    enteredVarValue.append(enteryValue)
    enteryValue.grid(row = 5 +  count, column = 1)
    count += 1

def clearAll(event):
    for i in range(0, len(enteredVarName)):
        enteredVarName[i].destroy()
        enteredVarValue[i].destroy()
        del enteredVarName[i]
        del enteredVarValue[i]
        count = 0
        print(enteredVarValue)




makeField = Button(text = 'Добавить значение')
makeField.bind('<Button-1>', add_line)


programStart = Button(text = 'Начать')
programStart.bind('<Button-1>', DocChenger)

clearButton = Button(text = 'Отчистить')
clearButton.bind('<Button-1>', clearAll)

upperText.grid(row = 0, column = 0)
makeField.grid(row = 7, column = 2)
enterFilename.grid(row = 1, column = 0)
programStart.grid(row = 8, column = 2)
clearButton.grid(row = 6, column = 2)

root.mainloop()
