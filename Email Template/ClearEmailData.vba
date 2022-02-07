Sub ClearEmailData()

'Total rows
totalRow = ThisWorkbook.Sheets("Email").Range("A3").End(xlDown).Row

If Sheets("Email").Range("A3").Value = "" Then

    MsgBox "Data is still empty"

Else

    With Sheets("Email")

        'Start removing data from row 3 to down
        .Range("A3:A" & totalRow).Value = ""
        .Range("B3:B" & totalRow).Value = ""
        .Range("C3:C" & totalRow).Value = ""
        .Range("M3:M" & totalRow).Value = ""
        .Range("N3:N" & totalRow).Value = ""
        .Range("J3:J" & totalRow).Value = ""
        .Range("K3:K" & totalRow).Value = ""
        .Range("L3:L" & totalRow).Value = ""
        
        'Start removing data from row 4 to down
        .Range("D4:D" & totalRow).Value = ""
        .Range("E4:E" & totalRow).Value = ""
        .Range("F4:F" & totalRow).Value = ""
        .Range("G4:G" & totalRow).Value = ""
        .Range("H4:H" & totalRow).Value = ""
        .Range("I4:I" & totalRow).Value = ""
        .Range("O4:O" & totalRow).Value = ""
        .Range("P4:P" & totalRow).Value = ""
        .Range("Q4:Q" & totalRow).Value = ""

    
    End With
End If