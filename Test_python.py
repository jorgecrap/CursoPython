
import openpyxl
import docx





# Excels

'''
excel_document = openpyxl.load_workbook('test.xlsx')
hojas = excel_document.get_sheet_names()
clientes = excel_document.get_sheet_by_name('Clientes')

celda = clientes.cell(row = 3, column = 2).value

rango = clientes['A1:B15']

for linea in rango:
	for celda in linea:
		print (celda.value)

'''


# Words

doc = docx.Document('demo.docx')
parrafos = doc.paragraphs

for par in parrafos:
	print (par.text)
	trozos = par.runs
	for tr in trozos:
		print('\t' + tr.text)