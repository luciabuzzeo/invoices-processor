# invoices-processor

It is based on a program, written in python language, which is called "Invoices Processor" and has the potential to autocomplete an excel file by extracting certain data from individual invoices in PDF files. 

I structured the solution in three parts: 

1) The extraction of all the text from PDF files to texts files: This part uses the Fitz library to open the original invoices in PDF file, then writes a text file for each PDF invoice it opens, and finally records if there are any invoices that could not be processed.

2) The extraction of the relevant data from the text files: This part opens the text file of all the invoices that were successfully processed, extracts the relevant lines, and stores them in a list. The relevant lines are those containing the name and ID number of the employee, the full amount paid, the valuation, the revenues paid on account, and the revenues on account passed on.

3) The insertion of the relevant data into an excel file: This part uses pandas to move the data from a python list to an excel file.  In the excel file, the data is in a table. The columns contain the names of the categories mentioned above (name, id number, full amount paid, valuation, revenues paid on account, revenues on account passed on).

