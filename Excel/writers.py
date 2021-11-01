from tkinter import *
"""
A printing function. Absolutely redundant
"""
def printer(string):
    print(string)


"""
A debug printing function with a logic switch
"""
def moro(c=0,debug=True):
    if(debug==True):
        print(c)

def write(*message, console, end="\n", sep=" "):
    text = ""
    for item in message:
        text += "{}".format(item)
        text += sep
    text += end
    console.insert(INSERT, text)
"""
A simple method for writing a column of values from a sheet to another
"""
def writeColumn(columnList,sheet,targetRow,targetCol,sourceRows):
    iterator = 0
    for i in range(sourceRows[0]-2,sourceRows[1]-1):
        sheet[targetCol+str(iterator+targetRow)] = columnList[i]
        iterator = iterator + 1

def clearColumn(sheet,targetRows,targetCol):
    for i in range(targetRows[0], targetRows[1]+1):
        sheet[targetCol + str(i)] = ""

# a method for forcing a single value for the entire column
def forceColumn(value,sheet,targetRow,targetCol,sourceRows):
    iterator = 0
    for i in range(sourceRows[0] - 2, sourceRows[1] - 1):
        sheet[targetCol + str(iterator + targetRow)] = value
        iterator = iterator + 1

#methods specific for the new products/uutuudet sheet

def writeUnitM(columnList,sheet,targetRow,targetCol,sourceRows):
    iterator = 0
    value=""
    for i in range(sourceRows[0]-2,sourceRows[1]-1):
        value=columnList[i]
        if (columnList[i]=="METER"):
            value = "M"
        sheet[targetCol+str(iterator+targetRow)] = value
        iterator = iterator + 1

def writeHyllytettava(columnList,sheet,targetRow,targetCol,sourceRows):
    iterator = 0
    value = ""
    for i in range(sourceRows[0] - 2, sourceRows[1] - 1):
        value = columnList[i]
        if (columnList[i] == 0):
            value = "Ei / No"
        elif(columnList[i] == 1):
            value = "Kyllä / Yes"
        sheet[targetCol + str(iterator + targetRow)] = value
        iterator = iterator + 1

def writePerusmaarayksikko(columnList,sheet,targetRow,targetCol,sourceRows):
    iterator = 0
    value = ""
    for i in range(sourceRows[0] - 2, sourceRows[1] - 1):
        value = columnList[i]
        if (value == "kpl" or value == "Rasia" or value == "pkt" or value == "Pullo" or value == "jm"):
            value = "KPL / PCS"
        elif (value == "jm"):
            value = "KG"
        sheet[targetCol + str(iterator + targetRow)] = value
        iterator = iterator + 1

def writeTuotemerkki(columnList,sheet,targetRow,targetCol,sourceRows):
    iterator = 0
    for i in range(sourceRows[0] - 2, sourceRows[1] - 1):
        sheet[targetCol + str(iterator + targetRow)] = columnList[i][0:1].upper()+columnList[i][1:]
        iterator = iterator + 1

def writeLuku(columnList,sheet,targetRow,targetCol,sourceRows):
    iterator = 0
    for i in range(sourceRows[0] - 2, sourceRows[1] - 1):
        value = columnList[i]
        if(isinstance(value,str)):
            sheet[targetCol + str(iterator + targetRow)] = float(value.replace(',' , '.'))
            iterator = iterator + 1
        else:
            sheet[targetCol + str(iterator + targetRow)] = value
            iterator = iterator + 1

def writePiikki(columnList,sheet,targetRow,targetCol,sourceRows):
    iterator = 0
    value = ""
    for i in range(sourceRows[0] - 2, sourceRows[1] - 1):
        value = str(columnList[i])
        if value == None:
            value = "Ei / No"
        elif "piikki" in value:
            value = "Kyllä / Yes"
        else:
            value = "Ei / No"
        sheet[targetCol + str(iterator + targetRow)] = value
        iterator = iterator + 1

