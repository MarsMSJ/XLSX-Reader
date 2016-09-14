# XLSX-Reader
Author Mario M Sarria Jr

Code for reading Microsoft Excel XLSX Files. The code goes row by row reading the data and outputs a list of row data. The purpose of this code is to extra data from Excel XLSX data only export files that generally come from crystal reports. 

Input: XLSX workbook file with a sheet full of data.
Output: RowList< ValueList<Row Values> >


The code reads the XLSX files and stores the values of the rows line by line. This code works best with XLSX workbook files that only have 1 sheet inside them. The code will work with workbooks that have multiple sheets moreover, you will get one List with rows and their values as if their was only one sheet. In other words, the code will not have any indicators as to which rows and their values correspond to which sheet. The output will consist of one sheet's row data followed by the next.
