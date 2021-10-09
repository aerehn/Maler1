from openpyxl import load_workbook
#from openpyxl import Workbook
import time
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