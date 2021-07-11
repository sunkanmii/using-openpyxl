import openpyxl

firstFile = r"C:\\Users\Sunkanmi-PC\Documents\newFile1.xlsx"
secondFile = r"C:\\Users\Sunkanmi-PC\Documents\newFile2.xlsx"

wbObj1 = openpyxl.load_workbook(firstFile.strip())    
wbObj2 = openpyxl.load_workbook(secondFile.strip()) 
wbk1 = wbObj1.worksheets[0]
wbk2 = wbObj2.worksheets[0]

# from the active attribute
sheetObj1 = wbObj1.active
sheetObj2 = wbObj2.active

hostNames = []
for myRow in range(2, sheetObj2.max_row):
    hostNames.append(sheetObj2.cell(row=myRow, column=1).value.lower())
wbObj2.close

for rowNum in range(2, sheetObj1.max_row):
    if wbk1.cell(row=rowNum, column=1).value.lower() in hostNames:
	    wbk1.cell(row=rowNum, column=2).value = "Yes"
    else:
	    wbk1.cell(row=rowNum, column=2).value = "Not on Sophos"
wbObj1.save(firstFile)
wbObj1.close