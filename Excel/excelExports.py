from openpyxl import load_workbook
from Excel.writers import printer, writeColumn, moro, write
from Excel.sheeters import writeUutuudet, writeVanhatTuotteet, writeNimet, writeToimitusyks
from tkinter import *
import pandas as pd
#import numpy as np
#from openpyxl import Workbook



# A master function that writes all the values to the target workbook

def ajaArvot(console,lahdeRivit=-1, kohdeRivi=32, lahde="export_SOK_taulukkovienti_2021-10-20_06-11-04.xlsx", kohde="SOK Käyttötavaroiden erätuotelomake (muut käyttötavarat) v2 33 (version 1) (version 1)_ROSTERi.xlsx",debug=True):




    moro("Loading source Workbook...", debug)
    lahdeDF = pd.read_excel(lahde)
    moro(lahdeDF,debug)
    moro("Loading target Workbook...", debug)
    SOK_Ke = load_workbook(filename=kohde)
    moro(SOK_Ke.sheetnames,debug)


    moro("Writing sheets...", debug)
    if lahdeRivit==-1:
        shape = lahdeDF.shape
        lahdeRivit=[2,shape[0]+1]
    #writeVanhatTuotteet(SOK_Ke, targetRow=kohdeRivi, source=lahdeDF,sourceRows = lahdeRivit)
    writeUutuudet(SOK_Ke, targetRow=kohdeRivi, source=lahdeDF,sourceRows = lahdeRivit)
    writeToimitusyks(SOK_Ke, targetRow=kohdeRivi, source=lahdeDF,sourceRows = lahdeRivit)
    writeNimet(SOK_Ke, targetRow=kohdeRivi, source=lahdeDF,sourceRows = lahdeRivit)
    write("Saving...",console=console)
    moro("Saving...",debug)
    try:
        SOK_Ke.save(
            filename=kohde)
    except PermissionError:
        moro("PermissionError. Kohdetiedosto todennäköisesti auki!")
        write("Lupavirhe! Kohdetiedosto todennäköisesti auki! Tallentaminen epäonnistui.",console=console)
    write("Saving process complete.",console=console)
    moro("Saving process complete.",debug)



#fuck you! we are in branch firstproto

