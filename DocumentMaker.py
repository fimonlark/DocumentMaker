from docxtpl import DocxTemplate
from tkinter import *
root = Tk()
root.title('docxChanger')
root.minsize(800, 400)
root.configure(background='#DBD7D6')
upperText = Label(fg='black', bg ='grey', width=30, height=1, text = 'Название шаблона').grid(row = 0, column = 0)
upperText1 = Label(fg='black', bg ='grey', width=30, height=1, text = 'Сохранить как').grid(row = 0, column = 1)
Label(fg='black', bg='grey', width=30, height=1, text='Имя переменной ').grid(column=0, row = 4)
Label(fg='black', bg='grey', width=30, height=1, text='Текст переменной ').grid(column=1, row = 4)
enterFilename = Entry(width = 50)
enterSaveName = Entry(width = 50)
count = 0
enteredVarName = []
enteredVarValue = []



def DocChanger(event):
    filename = str(enterFilename.get())
    doc = DocxTemplate(filename)
    context = {}
    for i in range(0, len(enteredVarName)):
        context[str(enteredVarName[i].get())] = str(enteredVarValue[i].get())

    doc.render(context)
    doc.save(str(enterSaveName.get() + '.docx'))

def add_line(event):
    global count
    print(count)
    enteryName = Entry()
    enteredVarName.append(enteryName)
    enteryName.grid(row = 5 + count)
    enteryValue = Entry()
    enteredVarValue.append(enteryValue)
    enteryValue.grid(row = 5 +  count, column = 1)
    count += 1

def clearAll(event):
    while enteredVarName != False:
        global count
        enteredVarName[0].destroy()
        enteredVarValue[0].destroy()
        del enteredVarName[0]
        del enteredVarValue[0]
        count = 0


makeField = Button(text = 'Добавить значение')
makeField.bind('<Button-1>', add_line)

programStart = Button(text = 'Начать')
programStart.bind('<Button-1>', DocChanger)

clearButton = Button(text = 'Отчистить все')
clearButton.bind('<Button-1>', clearAll)

makeField.grid(row = 7, column = 2)
enterFilename.grid(row = 1, column = 0)
enterSaveName.grid(row = 1, column = 1)
programStart.grid(row = 8, column = 2)
clearButton.grid(row = 6, column = 2)

root.mainloop()
