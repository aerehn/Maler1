from openpyxl import load_workbook
from Excel.writers import *
#from Excel.writers import printer, writeColumn, moro, write, forceColumn,writeUnitM, writeHyllytettava, writePerusmaarayksikko, writeTuotemerkki, writeLuku
def writeUutuudet(workbook, targetRows, source, sourceRows, console):
    debug = True
    offset = 0
    targetRow = targetRows[0]
    print(targetRow)
    print(type(targetRow))
    moro("Writing 1. Uutuudet - New articles",debug)
    sheet = workbook.get_sheet_by_name('1. Uutuudet - New articles')
    try:
        writeColumn(source['sku'], sheet, targetRow + offset, "AB", sourceRows=sourceRows)
        writeColumn(source['etiketin_lisateksti_25-fi_FI'], sheet, targetRow + offset, "EN", sourceRows=sourceRows)
        writeColumn(source['koko'], sheet, targetRow + offset, "BR", sourceRows=sourceRows)
        writeColumn(source['kplpaketissa'], sheet, targetRow + offset, "AO", sourceRows=sourceRows)
        writeColumn(source['myyntierana_hyllytettava'], sheet, targetRow + offset, "EC", sourceRows=sourceRows)
        writeColumn(source['raaka_aine_materiaali'], sheet, targetRow + offset, "BW", sourceRows=sourceRows)
        writeColumn(source['uusi_tilausnumero'], sheet, targetRow + offset, "AF", sourceRows=sourceRows)
        writeColumn(source['PKT_GTIN'], sheet, targetRow + offset, "AM", sourceRows=sourceRows)
        writeColumn(source['savy_vari-fi_FI'], sheet, targetRow + offset, "BT", sourceRows=sourceRows)
        writeColumn(source['tekninenvarinumero'], sheet, targetRow + offset, "BS", sourceRows=sourceRows)
        writeColumn(source['tilausnumero'], sheet, targetRow + offset, "AQ", sourceRows=sourceRows)
        writeColumn(source['tullikoodi_nimike'], sheet, targetRow + offset, "ET", sourceRows=sourceRows)
        writeColumn(source['tuotekuvaus_markkinointiteksti-fi_FI'], sheet, targetRow + offset, "EV", sourceRows=sourceRows)
        writeColumn(source['hyllynreuna_25-fi_FI'], sheet, targetRow + offset, "EL", sourceRows=sourceRows)
        writeColumn(source['hyllynreuna_25-sv_SE'], sheet, targetRow + offset, "EP", sourceRows=sourceRows)
        writeColumn(source['tuotenimi_40_merkkia-en_GB'], sheet, targetRow + offset, "R", sourceRows=sourceRows)
        writeColumn(source['tuotenimi_40_merkkia-fi_FI'], sheet, targetRow + offset, "P",
                    sourceRows=sourceRows)
        forceColumn("DD: suora/direct", sheet, targetRow + offset, "CP", sourceRows=sourceRows)
        forceColumn("Kyllä / Yes", sheet, targetRow + offset, "AS", sourceRows=sourceRows)
        forceColumn("EUR", sheet, targetRow + offset, "CU", sourceRows=sourceRows)
        forceColumn("FIN tax class 1: yleinen verokanta (only finnish suppliers)", sheet, targetRow + offset, "CX",
                    sourceRows=sourceRows)
        forceColumn("246: Suomi / Finland", sheet, targetRow + offset, "ER", sourceRows=sourceRows)
        writeUnitM(source['jmpaketissa-unit'], sheet, targetRow + offset, "AJ", sourceRows=sourceRows)
        writeHyllytettava(source["myyntierana_hyllytettava"], sheet, targetRow + offset, "EC", sourceRows=sourceRows)
        writePiikki(source["hyllytystapa"], sheet, targetRow + offset, "EE", sourceRows=sourceRows)
        forceColumn("Ei / No", sheet, targetRow + offset, "EG", sourceRows=sourceRows)
        writePerusmaarayksikko(source['tuotteen_perusmaarayksiko'], sheet, targetRow + offset, "X", sourceRows=sourceRows)
        writeTuotemerkki(source['tuotemerkki'], sheet, targetRow + offset, "T", sourceRows=sourceRows)
        writeLuku(source['pituus'], sheet, targetRow + offset, "DU", sourceRows=sourceRows)
        writeLuku(source['pituus'], sheet, targetRow + offset, "AH", sourceRows=sourceRows)
        writeLuku(source['leveys'], sheet, targetRow + offset, "DW", sourceRows=sourceRows)
        writeLuku(source['korkeus'], sheet, targetRow + offset, "DY", sourceRows=sourceRows)
        writeLuku(source['tuotteen_nettopaino'], sheet, targetRow + offset, "DQ", sourceRows=sourceRows)
        writeLuku(source['tuotteen_bruttopaino'], sheet, targetRow + offset, "DO", sourceRows=sourceRows)
    except KeyError as err:
        message1 = "Error: "+str(err)+"\n"
        message2 = ("Jos virheilmoituksen loppu on muotoa: \n"+
             "   Error: numero => Rivejä ei ole noin montaa!\n"+
             "   Error: 'Atribuutti' => vaadittavaa atribuuttia ei löydy lähdekansiosta\n"+
             "Uutuuksia ei kirjoitettu loppuun"
             )
        message = message1 + message2
        moro(message)
        write(message,  console=console)
    # KeyError = KeyError: numero Rivejä ei ole noin montaa!
    # KeyError = KeyError: 'Atribuutti' => vaadittavaa atribuuttia ei löydy lähdekansiosta
def writeVanhatTuotteet(workbook, targetRows, source, sourceRows, console):
    offset = 0
    targetRow = targetRows[0]
    sheet = workbook.get_sheet_by_name("Vanhat tuotteet - Old articles")
    #sku
    writeColumn(source['sku'], sheet, targetRow + offset, "A", sourceRows =sourceRows)
    #
    writeColumn(source['pitka_tuotenimi-fi_FI'], sheet, targetRow + offset, "C", sourceRows =sourceRows)

def writeToimitusyks(workbook, targetRows, source, sourceRows, console):
    offset = 0
    debug = True
    targetRow = targetRows[0]
    moro("Writing 2. Toimitusyks. -  Deliv. units", debug)
    sheet = workbook.get_sheet_by_name('2. Toimitusyks. -  Deliv. units')
    #writables get written
    try:
        writeLuku(source['korkeus'], sheet, targetRow + offset, "M", sourceRows=sourceRows)
        writeLuku(source['lavakorkeus'], sheet, targetRow + offset, "BV", sourceRows=sourceRows)
        writeLuku(source['lavanbruttopaino'], sheet, targetRow + offset, "BX", sourceRows=sourceRows)
        writeLuku(source['lavanleveys'], sheet, targetRow + offset, "BT", sourceRows=sourceRows)
        writeLuku(source['lavanpituus'], sheet, targetRow + offset, "BR", sourceRows=sourceRows)
        writeLuku(source['leveys'], sheet, targetRow + offset, "K", sourceRows=sourceRows)
        writeLuku(source['paketinbruttopaino'], sheet, targetRow + offset, "AR", sourceRows=sourceRows)
        writeLuku(source['paketti_korkeus'], sheet, targetRow + offset, "AP", sourceRows=sourceRows)
        writeLuku(source['paketti_leveys'], sheet, targetRow + offset, "AN", sourceRows=sourceRows)
        writeLuku(source['paketti_syvyys'], sheet, targetRow + offset, "AL", sourceRows=sourceRows)
        writeLuku(source['pituus'], sheet, targetRow + offset, "I", sourceRows =sourceRows)
        writeLuku(source['tuotteen_bruttopaino'], sheet, targetRow + offset, "O", sourceRows =sourceRows)
    except KeyError as err:
        message1 = "Error: "+str(err)+"\n"
        message2 = ("Jos virheilmoituksen loppu on muotoa: \n"+
             "   Error: numero => Rivejä ei ole noin montaa!\n"+
             "   Error: 'Atribuutti' => vaadittavaa atribuuttia ei löydy lähdekansiosta\n"+
             "Toimitusyksiköitä ei kirjoitettu loppuun"
             )
        message = message1 + message2
        moro(message)
        write(message,  console=console)

def writeNimet(workbook, targetRows, source, sourceRows, console):
    offset = 0
    debug=True
    targetRow=targetRows[0]
    moro("Writing 3. Nimet - Names", debug)
    sheet = workbook.get_sheet_by_name('3. Nimet - Names')
    try:
        writeColumn(source['pitka_tuotenimi-en_GB'], sheet, targetRow + offset, "U", sourceRows=sourceRows)
        writeColumn(source['pitka_tuotenimi-fi_FI'], sheet, targetRow + offset, "C", sourceRows=sourceRows)
        writeColumn(source['pitka_tuotenimi-sv_SE'], sheet, targetRow + offset, "M", sourceRows=sourceRows)
        writeColumn(source['tuotenimi_40_merkkia-sv_SE'], sheet, targetRow + offset, "O", sourceRows=sourceRows)
    except KeyError as err:
        message1 = "Error: "+str(err)+"\n"
        message2 = ("Jos virheilmoituksen loppu on muotoa: \n"+
             "   Error: numero => Rivejä ei ole noin montaa!\n"+
             "   Error: 'Atribuutti' => vaadittavaa atribuuttia ei löydy lähdekansiosta\n"+
             "Nimiä ei kirjoitettu loppuun"
             )
        message = message1+message2
        moro(message)
        write(message, console=console)