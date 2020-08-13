Attribute VB_Name = "Module3"
Sub GetExcelData()
Application.ScreenUpdating = False
Application.DisplayAlerts = False

'get network username for filepaths
Dim NetworkUN As Object
Set NetworkUN = CreateObject("WScript.Network")
UserName = NetworkUN.UserName & "." & NetworkUN.UserDomain

Dim accpath As String
Dim objAccess As Access.Application
Dim excelpath As String
Dim foldername As String
Dim fname As String

foldername = "C:\users\" & UserName & "path to folder"
        
        If Right(foldername, 1) <> Application.PathSeparator Then foldername = foldername & Application.PathSeparator
            fname = Dir(foldername & "*.xlsx")
            
    'begin looping through files here
Do While Len(fname)
    
With Workbooks.Open(foldername & fname)

Dim excelwb As Workbook
Set excelwb = ActiveWorkbook

'''Begin macro that will execute across all .xlsx files within the designated folder
excelwb.Worksheets(1).Range("A1").Value = "Hello World"

'''End macro here

'''enter the file path to the access database
accpath = "C:\users\" & UserName & "file path to access database file"
excelpath = foldername & fname

Set objAccess = New Access.Application
Call objAccess.OpenCurrentDatabase(accpath)
objAccess.Visible = False
Call objAccess.DoCmd.TransferSpreadsheet(acImport, _
acSpreadsheetTypeExcel12Xml, "Table1", excelpath, _
True, "A1:E" & lr) ''' "A1:E" & lr" is a range being copied fromt he excel sheet. Change to fit your needs.

End With

fname = Dir

excelwb.Save
excelwb.Close

Loop
MsgBox "Complete"

Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
