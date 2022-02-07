Sub CopyEmail()

'Copy from sheets "MOP", Columns "Relative NE" to sheets "Email", columns "DUID"
ThisWorkbook.Worksheets("MOP").ListObjects("MOP").ListColumns(1).DataBodyRange.Select
Selection.Copy
ThisWorkbook.Worksheets("Email").Range("A3").PasteSpecial Paste:=xlPasteValues
ThisWorkbook.Worksheets("Email").Range("A3").BorderAround

'Copy from sheets "MOP", Columns "Subcont" to sheets "Email", columns "Subcon"
ThisWorkbook.Worksheets("MOP").ListObjects("MOP").ListColumns(2).DataBodyRange.Select
Selection.Copy
ThisWorkbook.Worksheets("Email").Range("B3").PasteSpecial Paste:=xlPasteValues
ThisWorkbook.Worksheets("Email").Range("B3").BorderAround

'Copy from sheets "MOP", Columns "Scope" to sheets "Email", columns "MOP Scope"
ThisWorkbook.Worksheets("MOP").ListObjects("MOP").ListColumns(3).DataBodyRange.Select
Selection.Copy
ThisWorkbook.Worksheets("Email").Range("C3").PasteSpecial Paste:=xlPasteValues
ThisWorkbook.Worksheets("Email").Range("C3").BorderAround

'Copy from sheets "MOP", Columns "Dependency Qty" to sheets "Email", columns "Impact SIte"
ThisWorkbook.Worksheets("MOP").ListObjects("MOP").ListColumns(10).DataBodyRange.Select
Selection.Copy
ThisWorkbook.Worksheets("Email").Range("J3").PasteSpecial Paste:=xlPasteValues
ThisWorkbook.Worksheets("Email").Range("J3").BorderAround

'Copy from sheets "MOP", Columns "Dependency Qty" to sheets "Email", columns "Impact SIte"
ThisWorkbook.Worksheets("MOP").ListObjects("MOP").ListColumns(11).DataBodyRange.Select
Selection.Copy
ThisWorkbook.Worksheets("Email").Range("K3").PasteSpecial Paste:=xlPasteValues
ThisWorkbook.Worksheets("Email").Range("K3").BorderAround

ThisWorkbook.Worksheets("MOP").ListObjects("MOP").ListColumns(6).DataBodyRange.Select
Selection.Copy
ThisWorkbook.Worksheets("Email").Range("N3").PasteSpecial Paste:=xlPasteValues
ThisWorkbook.Worksheets("Email").Range("N3").BorderAround

ThisWorkbook.Worksheets("MOP").ListObjects("MOP").ListColumns(18).DataBodyRange.Select
Selection.Copy
ThisWorkbook.Worksheets("Email").Range("L3").PasteSpecial Paste:=xlPasteValues
ThisWorkbook.Worksheets("Email").Range("L3").BorderAround

End Sub
