Sub PrintList()

'Declare the var for new workbook
Dim new_wb As Workbook

'Create new blank workbook
Set new_wb = Workbooks.Add

'Get the location for saving the file (with the dialog-box)
save_loc = Application.GetSaveAsFilename(InitialFileName:="CR MW Plan", fileFIlter:="Excel Files (*.xlsx), *.xlsx")

'Open new blank workbook
new_wb.Activate

'Open the current workbook
ThisWorkbook.Activate

'Copy Table Header
ThisWorkbook.Worksheets("MOP").Range("A5:U5").Copy
new_wb.Worksheets("Sheet1").Range("A1").PasteSpecial Paste:=xlPasteFormats
ThisWorkbook.Worksheets("MOP").Range("A5:U5").Copy
new_wb.Worksheets("Sheet1").Range("A1").PasteSpecial Paste:=xlPasteValues
new_wb.Worksheets("Sheet1").Range("A1:U1").AutoFilter

'Copy Table Value
ThisWorkbook.Worksheets("MOP").Range("MOP").Select
Selection.Copy
new_wb.Worksheets("Sheet1").Range("A2").PasteSpecial Paste:=xlPasteFormats
ThisWorkbook.Worksheets("MOP").Range("MOP").Select
Selection.Copy
new_wb.Worksheets("Sheet1").Range("A2").PasteSpecial Paste:=xlPasteValues

'Remove the columns
new_wb.Worksheets("Sheet1").Columns("L").Delete 'Remove column "Implact Data Source"
new_wb.Worksheets("Sheet1").Columns("L").Delete 'Remove column "RFC Number" after "Impact Data Source"
new_wb.Worksheets("Sheet1").Columns("N").Delete 'Remove column "Originator"
new_wb.Worksheets("Sheet1").Columns("N").Delete 'Remove column "Ticket Status"
new_wb.Worksheets("Sheet1").Columns("O").Delete 'Remove column "Email File Name"
new_wb.Worksheets("Sheet1").Columns("P").Delete 'Remove column "ATF Number"

'Autofit the columns
new_wb.Worksheets("Sheet1").Columns("A:J").AutoFit
new_wb.Worksheets("Sheet1").Columns("L:O").AutoFit


'Proceed to save if it's not cancelled and close
If save_loc <> False Then
new_wb.SaveAs save_loc
new_wb.Close SaveChanges:=True

'Don't proceed to save if it's cancelled and close new workbook without saving it
Else
MsgBox "The file did not save.", vbCritical
new_wb.Close SaveChanges:=False
End If


End Sub
