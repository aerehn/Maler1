from openpyxl import load_workbook
# from openpyxl import Workbook
import time
from tkinter import *
from tkinter.filedialog import askopenfilename
from Excel.excelExports import ajaArvot
from Excel.writers import moro, write
debug = True
# root window/widget
root = Tk()
root.wm_attributes("-transparentcolor", 'grey')
root.title("Excel python projekti")
root.geometry("1181x665")
sourceExcel = ""
kohdeExcel = ""
lahdeRivit = [1, 1]
kohdeRivi = 1
pady=2
padx=2
#background photo
bg = PhotoImage(file = "NG58Qz1.png")
label1 = Label( root, image = bg)
label1.place(x = 0, y = 0)
#console

Console = Text(root, height=18)
Console.grid(row=1, column=3,sticky=W, pady=pady, padx=padx, rowspan=8)
"""
def write(*message, end = "\n", sep = " "):
    text = ""
    for item in message:
        text += "{}".format(item)
        text += sep
    text += end
    Console.insert(INSERT, text)
"""
# created a label(widget)
labelBg = "#c9c9c9"
myLabel = Label(root, text="Excel ohjelma", bg=labelBg)
lahdeLabel = Label(root, text="Lähde: -", justify=LEFT, bg=labelBg)
kohdeLabel = Label(root, text="Kohde: -", justify=LEFT, bg=labelBg)
# entries
lahdeEntry = Entry(root, width=10, bg="#a6a6a6", borderwidth=3)
kohdeEntry = Entry(root, width=10, bg="#a6a6a6", borderwidth=3)
# putting to screen
myLabel.grid(row=0, column=0,sticky=W, pady=pady, padx=padx)
lahdeLabel.grid(row=2, column=1,sticky=W, pady=pady, padx=padx)
kohdeLabel.grid(row=3, column=1,sticky=W, pady=pady, padx=padx)
lahdeEntry.grid(row=4, column=1,sticky=W, pady=pady, padx=padx)
kohdeEntry.grid(row=5, column=1,sticky=W, pady=pady, padx=padx)

# buttons
def lataaLahde():
    global sourceExcel
    filename = askopenfilename()
    sourceExcel = filename
    if(len(filename)>0):
        lahdeLabel.config(text="Lähde: " + sourceExcel)
    moro("lähde nappula"+ sourceExcel,debug)


lahdeButton = Button(root, text="Valitse lähdetiedosto", padx=10, pady=5, command=lataaLahde)
lahdeButton.grid(row=2, column=0,sticky=W, pady=pady, padx=padx)

def lataaKohde():
    global kohdeExcel
    filename = askopenfilename()
    kohdeExcel = filename
    if (len(filename) > 0):
        kohdeLabel.config(text="Kohde: " + kohdeExcel)
    moro("kohde nappula"+ kohdeExcel,debug)


kohdeButton = Button(root, text="Valitse kohdetiedosto", padx=10, pady=5, command=lataaKohde)
kohdeButton.grid(row=3, column=0,sticky=W, pady=pady, padx=padx)

def valitseLahdeRivi():
    global lahdeRivit
    row = lahdeEntry.get()
    moro(row,debug)
    # we check entry is a number
    numbers = []
    for word in row.split():
        if word.isdigit():
            numbers.append(int(word))
    if (len(numbers) == 2):
        if(numbers[0]<numbers[1]):
            lahdeRivit = numbers
            lahdeRiviButton.configure(text="Lähderivit: {} - {}".format(lahdeRivit[0], lahdeRivit[1]))


lahdeRiviButton = Button(root, text="Lähderivit", padx=10, pady=5, command=valitseLahdeRivi)
lahdeRiviButton.grid(row=4, column=0,sticky=W, pady=pady, padx=padx)

def valitseKohdeRivi():
    global kohdeRivi
    row = kohdeEntry.get()
    moro(row,debug)
    # we check entry is a number
    numbers = []
    for word in row.split():
        if word.isdigit():
            numbers.append(int(word))
    if len(numbers) == 1:
        kohdeRivi = numbers[0]
        kohdeRiviButton.configure(text="Kohderivi: {}".format(kohdeRivi))


kohdeRiviButton = Button(root, text="Kohderivi", padx=10, pady=5, command=valitseKohdeRivi)
kohdeRiviButton.grid(row=5, column=0,sticky=W, pady=pady, padx=padx)

def run():
    global runButton
    global Console
    if (len(sourceExcel)>0)&(len(kohdeExcel)>0)&(sourceExcel!=kohdeExcel)&(kohdeRivi>1)&(lahdeRivit[0]>1):
        moro(2,debug)
        write("ajetaan exceleitä", console=Console)
        ajaArvot(Console,lahdeRivit,kohdeRivi,sourceExcel,kohdeExcel,debug)
    else: #this is for testing only
        ajaArvot(Console, debug=debug)





runButton = Button(root, text="Aja", padx=10, pady=5, command=run)
runButton.grid(row=6, column=0,sticky=W, pady=pady, padx=padx)
root.wm_attributes("-transparentcolor", 'grey')
# main loop
root.mainloop()
