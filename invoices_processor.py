import fitz
import pandas as pd
import xlsxwriter
import openpyxl
from os import listdir,remove
from os.path import isfile, join
from datetime import datetime
import numpy as np


# Extraction of all the text from PDF files to texts files

def extract_text(invoice_filename, unsuccessfully_processed_names):
    invoice = fitz.open(invoice_filename)

    first_page = invoice.load_page(0)
    text = first_page.get_text("text")
    
    text_filename = invoice_filename.replace(".pdf",".txt")
    with open(text_filename, 'w') as f:
        f.write(text)
       
    with open(text_filename) as f:
        row_count = sum(1 for line in f)

    if row_count==TEMPLATE1["ROW_COUNT1"]:
        return 1
    elif row_count==TEMPLATE2["ROW_COUNT2"]:
        return 1
    elif row_count==TEMPLATE3["ROW_COUNT3"]:
        return 1
    elif row_count==TEMPLATE4["ROW_COUNT4"]:
        return 1
    else:
        unsuccessfully_processed_names.append(invoice_filename.lstrip('./invoices/').rstrip('.pdf'))
#        remove(text_filename)
        return 0

TEMPLATE1 = {
    "ROW_COUNT1": 119,
    "NAME1": 8,
    "NIF1": 7,
    "Importe_Integro_Satisfecho1": 18,
    "Valoracion1": 27, 
    "Ingresos_a_cuenta_efectuados1": 26,
    "Ingresos_a_cuenta_repercutidos1": 25,
    }
    
TEMPLATE2 = {
    "ROW_COUNT2": 50, 
    "NAME2": 8,
    "NIF2": 7,
    "Importe_Integro_Satisfecho2": 19, 
    }

TEMPLATE3 = {
    "ROW_COUNT3": 12,
    "NAME3": 3,
    "NIF3": 4,
    "Importe_Integro_Satisfecho3": 5,
    "Valoracion3": 7, 
    "Ingresos_a_cuenta_efectuados3": 8,
    "Ingresos_a_cuenta_repercutidos3": 9,
    }
    
TEMPLATE4 = {
    "ROW_COUNT4": 13, 
    "NAME4": 3,
    "NIF4": 4,
    "Importe_Integro_Satisfecho4": 5, 
    "Valoracion4": 7, 
    "Ingresos_a_cuenta_efectuados4": 8,
    "Ingresos_a_cuenta_repercutidos4": 9,
    }

successfully_processed_count = 0
unsuccessfully_processed_names = []
working_directory = "./invoices"
invoices = [f for f in listdir(working_directory) if isfile(join(working_directory, f))]


for invoice in invoices: 
    if invoice.endswith(".pdf"):
        successfully_processed_count += extract_text(working_directory+"/"+invoice, unsuccessfully_processed_names)
           
successfully_processed_invoices = [f for f in listdir(working_directory) if isfile(join(working_directory, f))]


# Extraction of the relevant data from the text files

excel_sheet=[]
#excel_sheet_lineF1=[]
current_datetime = datetime.now()

for successfully_processed_invoice in successfully_processed_invoices:    
        if successfully_processed_invoice.endswith(".txt"):     
            data = open(working_directory+"/"+successfully_processed_invoice)
            with open(working_directory+"/"+successfully_processed_invoice) as f:
                row_count = sum(1 for line in f)
            M=[]
            excel_sheet_line=[]     
            if row_count==TEMPLATE1["ROW_COUNT1"]:
                lines_to_read = [TEMPLATE1["NAME1"], TEMPLATE1["NIF1"], TEMPLATE1["Importe_Integro_Satisfecho1"], TEMPLATE1["Valoracion1"], TEMPLATE1["Ingresos_a_cuenta_efectuados1"], TEMPLATE1["Ingresos_a_cuenta_repercutidos1"]]          
                for position, line in enumerate(data):
                    if position in lines_to_read: 
                        M.append(line.rstrip('\n'))
                        excel_sheet_line=([successfully_processed_invoice.rstrip('.txt')])
                        for element in M:
                            excel_sheet_line.append(element)      
                nif = excel_sheet_line[1]
                name = excel_sheet_line[2]
                valoracion = excel_sheet_line[4]
                ingresos_a_cuenta_repercutidos = excel_sheet_line[6]
                excel_sheet_line[1] = name
                excel_sheet_line[2] = nif
                excel_sheet_line[6] = valoracion
                excel_sheet_line[4] = ingresos_a_cuenta_repercutidos
                excel_sheet.append(excel_sheet_line)
                

            elif row_count==TEMPLATE2["ROW_COUNT2"]:
                lines_to_read = [TEMPLATE2["NAME2"], TEMPLATE2["NIF2"], TEMPLATE2["Importe_Integro_Satisfecho2"]]          
                for position, line in enumerate(data):
                    if position in lines_to_read: 
                        M.append(line.rstrip('\n'))
                        excel_sheet_line=([successfully_processed_invoice.rstrip('.txt')])
                        for element in M:
                            excel_sheet_line.append(element)  
                nif = excel_sheet_line[1]
                name = excel_sheet_line[2]
                excel_sheet_line[1] = name
                excel_sheet_line[2] = nif
                excel_sheet.append(excel_sheet_line)

                
            elif row_count==TEMPLATE3["ROW_COUNT3"]:
                lines_to_read = [TEMPLATE3["NAME3"], TEMPLATE3["NIF3"], TEMPLATE3["Importe_Integro_Satisfecho3"], TEMPLATE3["Valoracion3"], TEMPLATE3["Ingresos_a_cuenta_efectuados3"], TEMPLATE3["Ingresos_a_cuenta_repercutidos3"]]          
                for position, line in enumerate(data):
                    if position in lines_to_read: 
                        M.append(line.rstrip('\n'))
                        excel_sheet_line=[successfully_processed_invoice.rstrip('.txt')]
                        for element in M:
                            excel_sheet_line.append(element)      
                excel_sheet.append(excel_sheet_line)


              
            elif row_count==TEMPLATE4["ROW_COUNT4"]:
                lines_to_read = [TEMPLATE4["NAME4"], TEMPLATE4["NIF4"], TEMPLATE4["Importe_Integro_Satisfecho4"], TEMPLATE4["Valoracion4"], TEMPLATE4["Ingresos_a_cuenta_efectuados4"], TEMPLATE4["Ingresos_a_cuenta_repercutidos4"]]          
                for position, line in enumerate(data):
                    if position in lines_to_read: 
                        M.append(line.rstrip('\n'))
                        excel_sheet_line=[successfully_processed_invoice.rstrip('.txt')]
                        for element in M:
                            excel_sheet_line.append(element)  
                excel_sheet.append(excel_sheet_line)


#  Insertion of the relevant data into an excel file    

str_current_datetime = str(current_datetime).replace(":","-")
df = pd.DataFrame(excel_sheet, columns=["Nombre documento", "Nombre" , "NIF", "Importe_Integro_Satisfecho", "Valoracion", "Ingresos_a_cuenta_efectuados", "Ingresos_a_cuenta_repercutidos"])
writer = pd.ExcelWriter(str_current_datetime+'.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='facturas personal', index=False)
writer.save()

print(f'Se procesaron correctamente { successfully_processed_count } documentos')
if len(unsuccessfully_processed_names) != 0:
    print("La lista de los documentos que no se pudieron procesar es: ")
    print(unsuccessfully_processed_names)

