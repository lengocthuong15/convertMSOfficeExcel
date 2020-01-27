# convertMSOfficeExcel
Simple VBS script for convert between MSOffice Excel file formats.
It is very useful when you want to convert a bunch of files.

How to use.

Before using, open the script, goto line 14 and adjust the file type that you want to convert
The default is converting from XLSX into XLS
Please ref this link for the filetypes: https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlfileformat-enumeration-excel

1. Open cmd
2. cscript msOfficeConvert_excel.vbs input_folder output_folder

Ex: 
Your XLSX files is in c:/temp/xlsx
You want convert XLSX files into XLS at: c:/temp/xls

cscript msOfficeConvert_excel.vbs  c:/temp/xlsx c:/temp/xls



