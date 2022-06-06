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

def donloade(event):
    try:
        files = open('vullist.xlsx', "wb")
        url = 'https://bdu.fstec.ru/files/documents/vullist.xlsx'
        headers = {
        'User-Agent': 'My User Agent 1.0',
        }
        response = requests.get(url, headers=headers)
        files.write(response.content)
        files.close()
    except requests.exceptions.ConnectionError:
        messagebox.showerror('Ошибка',
                             'Неудалось загрузить файл, проверьте интернет подключение')
def dateOn(event):  #По дате
    labelDateInfo.configure(state=NORMAL)
    textBoxFromDate.configure(state=NORMAL)
    textBoxToDate.configure(state=NORMAL)
    labelFromDate.configure(state=NORMAL)
    labelToDate.configure(state=NORMAL)
def dateOff(event):  #Без даты
    textBoxFromDate.delete(0, END)
    textBoxToDate.delete(0, END)
    labelDateInfo.configure(state=DISABLED)
    textBoxFromDate.configure(state=DISABLED)
    textBoxToDate.configure(state=DISABLED)
    labelFromDate.configure(state=DISABLED)
    labelToDate.configure(state=DISABLED)
def AnalysysWithDate(event):  # Функция для проверки правильности ввода даты
    chrb = radioButtonDateVar.get()  # Заносим в переменную chrb состояние радиокнопок (1 или 0)
    if chrb == 0:  # Если радиокнопка "По дате" включена (1)
        Analysys(event)  # Выполняем функцию Analysys
    else:
        dataFrom = textBoxFromDate.get()
        dataTo = textBoxToDate.get()
        if len(dataFrom and dataTo) == 10 and (dataFrom[2] and dataTo[2]) == '.' and (dataFrom[5] and dataTo[5]) == '.' and dataFrom[6:].isnumeric() and dataTo[6:].isnumeric() and dataFrom[:2].isnumeric() and dataTo[:2].isnumeric() and dataFrom[3:5].isnumeric() and dataTo[3:5].isnumeric():
          tsFrom = datetime(year=int(dataFrom[6:]), month=int(dataFrom[3:5]), day=int(dataFrom[:2]))
          tsTo = datetime(year=int(dataFrom[6:]), month=int(dataFrom[3:5]), day=int(dataFrom[:2]))
          if (tsFrom.date and tsTo.day) > 0 and (tsFrom.day and tsTo.day) < 32 and (tsFrom.month and tsTo.month) > 0 and (tsFrom.month and tsTo.month) < 13 and (tsFrom.year and tsTo.year) > 1900:
            dataFrom = datetime.strptime(textBoxFromDate.get(), "%d.%m.%Y")
            dataTo = datetime.strptime(textBoxToDate.get(), "%d.%m.%Y")
            Analysys(event)
          else:
            messagebox.showerror('Ошибка',
                                 'Некорректно введена дата')  # Если дата введена некорректно - выводим окно с ошибкой
        else:
          messagebox.showerror('Ошибка',
                               'Некорректно введена дата')  # Если дата введена некорректно - выводим окно с ошибкой


def Analysys(event):  # Функция поиска уязвимостей

    try:
        try:
            workbook = xlrd2.open_workbook('vullist.xlsx')  # открываем книгу xlsx

            sheet = workbook.sheet_by_index(0)  # получаем доступ к первой странице книги (нумерация страниц с 0)

            row = sheet.nrows  # определяем количество записей (строк) на листе
            if row != 0:
                names = sheet.col_values(4)
                status = sheet.col_values(14)

                danger_lavels = sheet.col_values(12)
                chrb = radioButtonDateVar.get()
                ddd = sheet.col_values(9)
                DDD = var1.get()
                sow=combo.get()
                global danger_low, danger_middle, danger_hight, danger_super
                danger_super, danger_hight, danger_middle, danger_low = 0, 0, 0, 0
                if chrb == 0: # Если радиокнопка "По дате" выключена (0)
                    dataFrom = datetime.strptime('01.01.1900',"%d.%m.%Y")
                    dataTo = datetime.strptime('17.06.3021', "%d.%m.%Y")
                else:
                    dataFrom = datetime.strptime(textBoxFromDate.get(), "%d.%m.%Y")
                    dataTo = datetime.strptime(textBoxToDate.get(), "%d.%m.%Y")

                for i in range(9, row):
                    if ddd[i] != '':
                        ddd[i] = datetime.strptime(ddd[i], "%d.%m.%Y")
                    else:
                        ddd[i] = datetime.strptime('01.01.1999', "%d.%m.%Y")
                for i in range(4, row):
                    if (str(ddd[i]) >= str(dataFrom)) and (str(ddd[i]) <= str(dataTo)):
                        if DDD == 1:
                            if status[i].find("Потенциальная уязвимость") >= 0 and names[i].find(sow) >= 0:
                                if danger_lavels[i][0] == 'К':  # Критический
                                    danger_super += 1

                                elif danger_lavels[i][0] == 'В':  # Высокий
                                    danger_hight += 1

                                elif danger_lavels[i][0] == 'С':  # Средний
                                    danger_middle += 1

                                else:  # Низкий
                                    danger_low += 1
                        else:
                            if names[i].find(sow) >= 0:

                              if danger_lavels[i][0] == 'К':  # Критический
                                danger_super += 1

                              elif danger_lavels[i][0] == 'В':  # Высокий
                                danger_hight += 1

                              elif danger_lavels[i][0] == 'С':  # Средний
                                danger_middle += 1

                              else:  # Низкий
                                danger_low += 1

                labelLowOut['text'] = danger_low
                labelMidOut['text'] = danger_middle
                labelHighOut['text'] = danger_hight
                labelSuperOut['text'] = danger_super
            else:
                messagebox.showerror('Ошибка',
                                     'Для анализа необходимо обновить базу уязвимостей!')
        except FileNotFoundError:
            messagebox.showerror('Ошибка',
                                 'Для анализа необходимо загрузить (Обновить) базу уязвимостей!')
    except xlrd2.biffh.XLRDError:
        messagebox.showerror('Ошибка',
                             'Для анализа необходимо загрузить (Обновить) базу уязвимостей!')

def diagramma(event):
    try:
        if danger_low == 0 and danger_middle == 0 and danger_hight == 0 and danger_super ==0:
            messagebox.showerror('Ошибка',
                                 'Для вывода диаграммы необходимо хотя бы одно значение отличное от 0!')
        else:
            labels = 'Низкий', 'Средний', 'Высокий', 'Критический'
            sizes = [danger_low, danger_middle, danger_hight, danger_super]

            colors = ("grey", "yellow", "orange",
            "red")
            fig1, ax1 = plt.subplots()
            explode = (0, 0.1, 0, 0)

            ax1.pie(sizes, wedgeprops=dict(width=1), colors=colors, explode=explode, labels=labels,autopct='%1.1f%%', shadow=True, startangle=90)
            patches, texts, auto = ax1.pie(sizes, wedgeprops=dict(width=1), colors=colors, shadow=True, startangle=90, explode=explode, autopct='%1.1f%%' )

            plt.legend(patches, labels, loc="best")
            okno=Tk()
            okno.title("Диаграмма уязвимостей")
            okno.configure(background='#a8e4a0')
            canvas = FigureCanvasTkAgg(fig1, master=okno)
            canvas.get_tk_widget().pack()
            canvas.draw()
    except NameError:
        messagebox.showerror('Ошибка',
                             'Для вывода диаграммы необходимо провести анализ!')

def Clear(event): #очистка полей
    labelLowOut['text'] = ""
    labelMidOut['text'] = ""
    labelHighOut['text'] = ""
    labelSuperOut['text'] = ""
    textBoxFromDate.delete(0, END)
    textBoxToDate.delete(0, END)


def SaveDocx(event):  # Функция для сохранения результатов в docx
  if labelLowOut['text'] == "" and labelMidOut['text'] == "" and labelHighOut['text'] == "" and labelSuperOut['text'] == "":
      messagebox.showerror('Ошибка',
                           'Для вывода отчёта в документ необходимо провести анализ!')
  else:
      if scale_widget.get() == 0:
        document = docx.Document()
        document.add_heading(combo.get(), 0)
        document.add_heading('Количество уязвимостей по уровням опасности', level=1)
        table = document.add_table(rows=4, cols=3)
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = '1'
        hdr_cells[1].text = 'Низкий'
        hdr_cells[2].text = str(labelLowOut['text'])
        hdr_cells1 = table.rows[1].cells
        hdr_cells1[0].text = '2'
        hdr_cells1[1].text = 'Средний'
        hdr_cells1[2].text = str(labelMidOut['text'])
        hdr_cells2 = table.rows[2].cells
        hdr_cells2[0].text = '3'
        hdr_cells2[1].text = 'Высокий'
        hdr_cells2[2].text = str(labelHighOut['text'])
        hdr_cells3 = table.rows[3].cells
        hdr_cells3[0].text = '4'
        hdr_cells3[1].text = 'Критический'
        hdr_cells3[2].text = str(labelSuperOut['text'])
        document.save('Анализ уязвимостей ' + combo.get() + '.docx')
      else:
        my_wb = openpyxl.Workbook()
        my_sheet = my_wb.active
        c1 = my_sheet.cell(row=1, column=1)
        c1.value = "Угроза"
        c3 = my_sheet.cell(row=2, column=1)
        c3.value = 'Низкая'
        c4 = my_sheet.cell(row=3, column=1)
        c4.value = 'Средняя'
        c5 = my_sheet.cell(row=4, column=1)
        c5.value = 'Высокая'
        c6 = my_sheet.cell(row=5, column=1)
        c6.value = 'Критическая'
        c2 = my_sheet.cell(row=1, column=2)
        c2.value = "Количество угроз"
        c7 = my_sheet.cell(row=5, column=2)
        c7.value = labelSuperOut['text']
        c8 = my_sheet.cell(row=2, column=2)
        c8.value = labelLowOut['text']
        c9 = my_sheet.cell(row=3, column=2)
        c9.value = labelMidOut['text']
        c10 = my_sheet.cell(row=4, column=2)
        c10.value = labelHighOut['text']
        my_wb.save('Анализ уязвимостей ' + combo.get() + '.xlsx')

def example1(event):
    def print_sel():
        textBoxFromDate.delete(0, END)
        DATA = str(cal.selection_get())
        data = DATA[8] + DATA[9] + "." + DATA[5] + DATA[6] + "." + DATA[0] + DATA[1] + DATA[2] + DATA[3]
        textBoxFromDate.insert(0, data)
        okno.update()

    top = tk.Toplevel(okno)

    cal = Calendar(top,
                   font="Arial 14", selectmode='day',
                   cursor="hand1", year=2018, month=2, day=5)
    cal.pack(fill="both", expand=True)
    ttk.Button(top, text="ok", command=print_sel).pack()

def example2(event):
    def print_sel():
        textBoxToDate.delete(0, END)
        DATA = str(cal.selection_get())
        data = DATA[8] + DATA[9] + "." + DATA[5] + DATA[6] + "." + DATA[0] + DATA[1] + DATA[2] + DATA[3]
        textBoxToDate.insert(0, data)
        okno.update()


    top = tk.Toplevel(okno)

    cal = Calendar(top,
                   font="Arial 14", selectmode='day',
                   cursor="hand1", year=2018, month=2, day=5)
    cal.pack(fill="both", expand=True)
    ttk.Button(top, text="ok", command=print_sel).pack()

okno = Tk() # Главное меню
okno.resizable(False, False)
okno.title("Анализ уязвимости")  # Название окна
okno.geometry("1000x600")  # Размер окна
okno.configure(background='#3034ab')

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

radioButtonDateVar = BooleanVar()  #кнопоки
radioButtonDateVar.set(0)
radioButtonDateOn = Radiobutton(okno, text="По дате", bg='#afdafc', variable=radioButtonDateVar, value=1)
radioButtonDateOn.bind('<Button-1>', dateOn)
radioButtonDateOn.pack(anchor=W)
radioButtonDateOn.place(x=0, y=380)

radioButtonDateOff = Radiobutton(okno, text="За все время", bg='#afdafc', variable=radioButtonDateVar, value=0)
radioButtonDateOff.bind('<Button-1>', dateOff)
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
buttonobnow.bind('<Button-1>', donloade)
buttonobnow.place(x=300, y=515)

buttonAnalysis = Button(okno, bg='#027ef2', font='Times 12', text="Анализ базы", width=13, height=2)
buttonAnalysis.place(x=450, y=515)
buttonAnalysis.bind('<Button-1>', AnalysysWithDate)

buttonClear = Button(okno, bg='#027ef2', font='Times 12', text="Удалить", width=13, height=2)
buttonClear.place(x=600, y=515)
buttonClear.bind('<Button-1>', Clear)

buttonSave = Button(okno, bg='#027ef2', font='Times 12', text="Сохранить", width=13, height=2)
buttonSave.place(x=750, y=515)
buttonSave.bind('<Button-1>', SaveDocx)

buttonDiagram = Button(okno, bg='#7fc7ff', font='Times 12', text="диаграмма", height=2)
buttonDiagram.place(x=190, y=460)
buttonDiagram.bind('<Button-1>', diagramma)

labelDate = Label(okno, text="Введите дату:", state=DISABLED, bg='#addfad', font='Times 13', fg='#000', width=30)
labelDate.place(x=200, y=380)

labelFromDate = Button(okno, text=" От:", state=DISABLED, bg='#ffd1d1', fg='black', width=5)
labelFromDate.place(x=200, y=410)
labelFromDate.bind('<Button-1>', example1)

textBoxFromDate=Entry(okno, state=DISABLED, width=15)
textBoxFromDate.place(x=250, y=410)

labelToDate = Button(okno, text="До:", state=DISABLED, bg='#ffd1d1', fg='black', width=5)
labelToDate.place(x=365, y=410)
labelToDate.bind('<Button-1>', example2)

textBoxToDate = Entry(okno, state=DISABLED, width=15)
textBoxToDate.place(x=415, y=410)

labelDateInfo = Label(okno, text="Анализ уязвимостей", bg='#78b6f0', font='Times 20', fg='#0f0901', width=50)
labelDateInfo.pack()

labelToInfo = Label(okno, bg='#78b6f0', fg='black', width=20)

#scale_widget = Scale(master=okno, from_=Word, to=Excel, orient="horizontal")
#scale_widget.place(x=750, y=420)
#(x=750, y=515)

label0 = Label(okno, text="Word")
label0.place(x=750, y=405)

label1 = Label(okno, text="Excel")
label1.place(x=820, y=405)
okno.mainloop()
