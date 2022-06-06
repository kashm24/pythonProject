import xlrd2
from tkinter import *
from tkinter import messagebox
from datetime import datetime
from tkinter import ttk
import tkinter as tk
import requests
import docx
import openpyxl
from tkinter.ttk import Combobox
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkcalendar import Calendar

okno = Tk() # Главное окно

okno.resizable(False, False)
okno.title("Анализ уязвимости")  # Название окна
okno.geometry("1000x600")  # Размер окна
okno.configure(background='#3034ab')
# Создание интерфейса

labelLow = Label(okno, width=13, height=2, bg='#fabbbf', font='Times 13', text="Низкий")
labelLow.place(x=0, y=80)

labelLowOut = Label(okno, bg='#99ffcc', font='Times 15', fg='black', width=5)
labelLowOut.place(x=0, y=125)

labelMid = Label(okno, width=13, height=2, bg='#f78d94', font='Times 13', text="Средний")
labelMid.place(x=0, y=150)

labelMidOut = Label(okno, bg='#99ffcc', font='Times 15', fg='black', width=5)
labelMidOut.place(x=0, y=195)

labelHigh = Label(okno, width=13, height=2, bg='#f55d67', font='Times 13', text="Высокий")
labelHigh.place(x=0, y=220)

labelHighOut = Label(okno, bg='#99ffcc', font='Times 15', fg='black', width=5)
labelHighOut.place(x=0, y=265)

labelSuper = Label(okno, width=13, height=2, bg='#f22c39', font='Times 13', text="Критический")
labelSuper.place(x=0, y=295)

labelSuperOut = Label(okno, bg='#99ffcc', font='Times 15', fg='black', width=5)
labelSuperOut.place(x=0, y=340)

radioButtonDateVar = BooleanVar()  # Создание кнопок
radioButtonDateVar.set(0)
radioButtonDateOn = Radiobutton(okno, text="По дате", bg='#afdafc', variable=radioButtonDateVar, value=1)
radioButtonDateOn.bind('<Button-1>', )
radioButtonDateOn.pack(anchor=W)
radioButtonDateOn.place(x=0, y=380)

radioButtonDateOff = Radiobutton(okno, text="За все время", bg='#afdafc', variable=radioButtonDateVar, value=0)
radioButtonDateOff.bind('<Button-1>', )
radioButtonDateOff.pack(anchor=W)
radioButtonDateOff.place(x=0, y=410)

combo = Combobox(okno)
combo['values'] = ('CentOS', 'Red Hat Enterprise', 'Red Hat Enterprise Linux', 'Red Hat Inc.')
combo.current(0)  # установите вариант по умолчанию
combo.grid(column=0, row=0)
combo.place(x=230, y=38)

var1 = BooleanVar()
var1.set(0)
c1 = Checkbutton(okno, bg = '#afdafc', text="Потенциальная уязвимость", variable=var1, onvalue=1, offvalue=0)
c1.pack(anchor=W, padx=10)
c1.grid(column=0, row=0)
c1.place(x=210, y=60)

buttonobnow = Button(okno, bg='#027ef2', font='Times 12', text="Обновить базу", width=13, height=2)
buttonobnow.bind('<Button-1>', )
buttonobnow.place(x=300, y=515)

buttonAnalysis = Button(okno, bg='#027ef2', font='Times 12', text="Анализ", width=13, height=2)
buttonAnalysis.place(x=450, y=515)
buttonAnalysis.bind('<Button-1>', )

buttonClear = Button(okno, bg='#027ef2', font='Times 12', text="Удалить", width=13, height=2)
buttonClear.place(x=600, y=515)
buttonClear.bind('<Button-1>', )

buttonSave = Button(okno, bg='#027ef2', font='Times 12', text="Сохранить", width=13, height=2)
buttonSave.place(x=750, y=515)
buttonSave.bind('<Button-1>', )

buttonDiagram = Button(okno, bg='#7fc7ff', font='Times 12', text="Вывести диаграмму", height=2)
buttonDiagram.place(x=190, y=460)
buttonDiagram.bind('<Button-1>', )

labelDate = Label(okno, text="Введите необходимую дату:", state=DISABLED, bg='#addfad', font='Times 13', fg='#000', width=30)
labelDate.place(x=200, y=380)

labelFromDate = Button(okno, text=" От:", state=DISABLED, bg='#ffd1d1', fg='black', width=5)
labelFromDate.place(x=200, y=410)
labelFromDate.bind('<Button-1>', )

textBoxFromDate=Entry(okno, state=DISABLED, width=15)
textBoxFromDate.place(x=250, y=410)

labelToDate = Button(okno, text="До:", state=DISABLED, bg='#ffd1d1', fg='black', width=5)
labelToDate.place(x=365, y=410)
labelToDate.bind('<Button-1>', )

textBoxToDate = Entry(okno, state=DISABLED, width=15)
textBoxToDate.place(x=415, y=410)

labelDateInfo = Label(okno, text="Анализ уязвимостей", bg='#78b6f0', font='Times 20', fg='#0f0901', width=50)
labelDateInfo.pack()

labelToInfo = Label(okno, bg='#78b6f0', fg='black', width=20)

scale_widget = Scale(master=okno, from_=0, to=1, orient="horizontal")
scale_widget.place(x=250, y=150)

label0 = Label(okno, text="Word")
label0.place(x=320, y=150)

label1 = Label(okno, text="Excel")
label1.place(x=250, y=150)
okno.mainloop()