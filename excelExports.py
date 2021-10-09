from openpyxl import load_workbook
from tkinter import *
from openpyxl import Workbook
def moro(c=0,debug=True):
    if(debug==True):
        print("moro",c)

def ajaArvot(lahdeRivit, kohdeRivi, lahde, kohde,console,debug):
    def write(*message, end="\n", sep=" "):
        text = ""
        for item in message:
            text += "{}".format(item)
            text += sep
        text += end
        console.insert(INSERT, text)

    """
    start_time = time.time()
    sok_pim = load_workbook(filename="PIM - SOK -infoa.xlsx")
    SOK_Ke = load_workbook(filename="SOK Käyttötavaroiden erätuotelomake (muut käyttötavarat) v2 33 (version 1) (version 1)_ROSTERi.xlsx")
    print(sok_pim.sheetnames)
    print(SOK_Ke.sheetnames)
    print("Loading took %s seconds" % (time.time()-start_time))
    start_time = time.time()
    sheet2 = SOK_Ke.active
    sheet = sok_pim.active
    sheet['A15'] = "Well hello there!"
    sheet2['A40'] = 'moro'
    sheet2['A41'] = sheet['ER1'].value
    print("Modifications took %s seconds" % (time.time()-start_time))
    start_time = time.time()
    SOK_Ke.save(filename="SOK Käyttötavaroiden erätuotelomake (muut käyttötavarat) v2 33 (version 1) (version 1)_ROSTERi.xlsx")
    sok_pim.save(filename="PIM - SOK -infoa.xlsx")
    print("Saving took %s seconds" % (time.time()-start_time))
    """
    write(lahdeRivit)

    write(kohdeRivi)
    moro("ajetaan exceleitä", debug)
    sok_pim = load_workbook(filename=lahde)

    SOK_Ke = load_workbook(filename=kohde)

    write(sok_pim.sheetnames)
    moro(sok_pim.sheetnames, debug)
    write(SOK_Ke.sheetnames)
    moro(SOK_Ke.sheetnames, debug)
