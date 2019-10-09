Option Explicit

Sub WriteLine ( strLine )
    WScript.Stdout.WriteLine strLine
End Sub

' How to use
' Open cmd
' Type: cscript msOfficeConvert_excel.vbs input_folder output_folder


' Add what type you need to save as in here
' Please Ref to this link: https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlfileformat-enumeration-excel
Const xlExcel8 = 56 ' Save as XLS


Dim inputDirectory
Dim inputFolder
Dim inFiles
Dim outputFolder
Dim inputFile
Dim outputFile
Dim curFile
Dim objExcel
Dim objWorkbook
Dim objPrintOptions
Dim objFso
Dim curDir



If WScript.Arguments.Count <> 2 Then
    WriteLine "You need to specify 2 input folders."
    WScript.Quit
End If

Set objFso = CreateObject("Scripting.FileSystemObject")

curDir = objFso.GetAbsolutePathName(".")

Set inputFolder = objFSO.GetFolder(WScript.Arguments.Item(0))
Set outputFolder = objFSO.GetFolder(WScript.Arguments.Item(1)) 

Set inFiles = inputFolder.Files

Set objExcel = CreateObject( "Excel.Application" )

For Each curFile in inFiles

Set inputFile = curFile

If Not objFso.FileExists( inputFile ) Then
    WriteLine "Unable to find your input file " & inputFile
    WScript.Quit
End If

WriteLine inputFile

Set objWorkbook = objExcel.Workbooks.Open(inputFile,0, true, 5)
on error resume next
objWorkbook.CheckCompatibility = false
objWorkbook.SaveAs outputFolder & "\" & curFile.Name & ".xls", xlExcel8

objWorkbook.Close

Next

objExcel.Quit
