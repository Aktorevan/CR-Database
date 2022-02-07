Sub Autofill_Email()

'Total rows
totalRow = ThisWorkbook.Sheets("Email").Range("A3").End(xlDown).Row

If Sheets("Email").Range("A3").Value = "" Then

    MsgBox "No Data to Fill Out"

Else

    With Sheets("Email")

    'Copy down
    .Range("D3:D" & totalRow).FillDown
    .Range("E3:E" & totalRow).FillDown
    .Range("F3:F" & totalRow).FillDown
    .Range("G3:G" & totalRow).FillDown
    .Range("H3:H" & totalRow).FillDown
    .Range("I3:I" & totalRow).FillDown
    .Range("O3:O" & totalRow).FillDown
    .Range("P3:P" & totalRow).FillDown
    .Range("Q3:Q" & totalRow).FillDown
    
    'Border All
    .Range("A3:Q" & totalRow + 2).BorderAround

    End With

End If

End Sub
