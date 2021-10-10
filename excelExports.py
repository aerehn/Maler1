from openpyxl import load_workbook
from tkinter import *
import pandas as pd
import numpy as np
from openpyxl import Workbook
def moro(c=0,debug=True):
    if(debug==True):
        print(c)

def ajaArvot(lahdeRivit=[2,2], kohdeRivi=32, lahde="export_Tuotteiden_vienti_XLSX_2021-10-09_14-42-15.xlsx", kohde="SOK Käyttötavaroiden erätuotelomake (muut käyttötavarat) v2 33 (version 1) (version 1)_ROSTERi.xlsx",console=0,debug=True):
    def write(*message, end="\n", sep=" "):
        text = ""
        for item in message:
            text += "{}".format(item)
            text += sep
        text += end
        console.insert(INSERT, text)


    #function for inserting data in column form
    def writeColumn(columnList,sheet,targetRow,targetCol,sourceRows,maxLen=0):
        for i in range(sourceRows[0]-2,sourceRows[1]-1):
            if maxLen>0:
                sheet[targetCol+str(i+targetRow)] = columnList[i][:maxLen]
            else:
                sheet[targetCol + str(i + targetRow)] = columnList[i]
    def writeVanhatTuotteet(workbook, targetRow, source, sourceRows):
        offset = 0
        sheet = workbook.get_sheet_by_name("Vanhat tuotteet - Old articles")
        #sku
        writeColumn(source['sku'], sheet, targetRow + offset, "A", sourceRows =sourceRows)
        #
        writeColumn(source['pitka_tuotenimi-fi_FI'], sheet, targetRow + offset, "C", sourceRows =sourceRows)

    #offset=32
    def writeUutuudet(workbook, targetRow, source, sourceRows):
        offset = 0

        sheet = workbook.get_sheet_by_name('1. Uutuudet - New articles')
        sourceCols = ['sku','pitka_tuotenimi-fi_FI','pitka_tuotenimi-fi_FI','tuotemerkki','kappalettalavalla','tuotteen_nettopaino','pituus','leveys']
        targetCols = ["N","O","P","T","CJ","DQ","DU","DW"]
        writeColumn(source['sku'], sheet, targetRow + offset, "N", sourceRows =sourceRows)
        writeColumn(source['sku'], sheet, targetRow + offset, "AB", sourceRows =sourceRows)
        writeColumn(source['tilausnumero'], sheet, targetRow + offset, "AQ", sourceRows =sourceRows)
        writeColumn(source['myyntieran_sisalto'], sheet, targetRow + offset, "AO", sourceRows =sourceRows)
        writeColumn(source['pitka_tuotenimi-fi_FI'], sheet, targetRow + offset, "O", sourceRows =sourceRows)
        writeColumn(source['pitka_tuotenimi-fi_FI'], sheet, targetRow + offset, "P", maxLen=40, sourceRows =sourceRows)
        #writeColumn(source['tuotenimi_40_merkkia-fi_FI'], sheet, row+offset, "P",sourceRows =sourceRows)
        writeColumn(source['tuotemerkki'], sheet, targetRow + offset, "T", sourceRows =sourceRows)
        writeColumn(source['kappalettalavalla'], sheet, targetRow + offset, "CJ", sourceRows =sourceRows)
        writeColumn(source['tuotteen_nettopaino'], sheet, targetRow + offset, "DQ", sourceRows =sourceRows)
        writeColumn(source['tuotteen_bruttopaino'], sheet, targetRow + offset, "DO", sourceRows =sourceRows)
        writeColumn(source['tuotekuvaus_markkinointiteksti-fi_FI'], sheet, targetRow + offset, "EV", sourceRows =sourceRows)
        writeColumn(source['raaka_aine_materiaali'], sheet, targetRow + offset, "BW", sourceRows =sourceRows)
        writeColumn(source['savy_vari-fi_FI'], sheet, targetRow + offset, "BT", sourceRows =sourceRows)
        writeColumn(source['koko'], sheet, targetRow + offset, "BR", sourceRows =sourceRows)
        writeColumn(source['pituus'], sheet, targetRow + offset, "DU", sourceRows =sourceRows)
        writeColumn(source['leveys'], sheet, targetRow + offset, "DW", sourceRows =sourceRows)
        writeColumn(source['korkeus'], sheet, targetRow + offset, "DY", sourceRows =sourceRows)
        sheet['AS39']="Kyllä / Yes"
    def writeToimitusyks(workbook, targetRow, source, sourceRows):
        offset = 0
        sheet = workbook.get_sheet_by_name('2. Toimitusyks. -  Deliv. units')
        #writeColumn(source['sku'], sheet, row + offset, "A")
        #writeColumn(source['pitka_tuotenimi-fi_FI'], sheet, row + offset, "B")
        #writeColumn(source['tilausnumero'], sheet, row + offset, "F")
        #writeColumn(source['kappalettalavalla'], sheet, row + offset, "BM")
        writeColumn(source['myyntieran_sisalto'], sheet, targetRow + offset, "H", sourceRows =sourceRows)
        writeColumn(source['myyntieran_materiaalin_paino'], sheet, targetRow + offset, "O", sourceRows =sourceRows)
        writeColumn(source['pituus'], sheet, targetRow + offset, "I", sourceRows =sourceRows)
        writeColumn(source['leveys'], sheet, targetRow + offset, "K", sourceRows =sourceRows)
        writeColumn(source['korkeus'], sheet, targetRow + offset, "M", sourceRows =sourceRows)
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

    moro("Loading source Workbook...", debug)
    lahdeDF = pd.read_excel(lahde)
    moro(lahdeDF,debug)
    moro("Loading target Workbook...", debug)
    SOK_Ke = load_workbook(filename=kohde)
    moro(SOK_Ke.sheetnames,debug)

    sheet2 = SOK_Ke.active
    moro("Writing sheets...", debug)
    #writeVanhatTuotteet(SOK_Ke, targetRow=kohdeRivi, source=lahdeDF,sourceRows = lahdeRivit)
    writeUutuudet(SOK_Ke, targetRow=kohdeRivi, source=lahdeDF,sourceRows = lahdeRivit)
    writeToimitusyks(SOK_Ke, targetRow=kohdeRivi, source=lahdeDF,sourceRows = lahdeRivit)
    write("Saving...")
    moro("Saving...",debug)
    try:
        SOK_Ke.save(
            filename="SOK Käyttötavaroiden erätuotelomake (muut käyttötavarat) v2 33 (version 1) (version 1)_ROSTERi.xlsx")
    except PermissionError:
        moro("PermissionError. Kohdetiedosto todennäköisesti auki!")
        write("Lupavirhe! Kohdetiedosto todennäköisesti auki! Tallentaminen epäonnistui.")
    write("Saving process complete.")
    moro("Saving process complete.",debug)



#fuck you! we are in branch firstproto

