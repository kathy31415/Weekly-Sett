Attribute VB_Name = "WeeklySettlement_NI_"
Sub WeeklySettlement_NI()

Dim CellArray() As Variant

'----------------------------------------------------SUMMARY SHEET

CellArray = Array("L17", "M17", "N17", "O17", "P17")

For Each cel In CellArray
    With Worksheets("Summary").Range(cel)
        .NumberFormat = "0.0"
        .Value = .Value
    End With
Next

'----------------------------------------------------REPORT SHEET
CellArray = Array("I:I", "M:M", "S:S", "V:V")

For Each cel In CellArray
    With Worksheets("REPORT").Range(cel)
        .NumberFormat = "0"
        .Value = .Value
    End With
Next

CellArray = Array("M:M", "S:S")

For Each cel In CellArray
    With Worksheets("REPORT").Range(cel)
        .NumberFormat = "0.00000"
        .Value = .Value
    End With
Next

'----------------------------------------------------STATEMENT SHEET

CellArray = Array("I:I")

For Each cel In CellArray
    With Worksheets("STATEMENT").Range(cel)
        .NumberFormat = "0"
        .Value = .Value
    End With
Next

CellArray = Array("L:L", "P:P")

For Each cel In CellArray
    With Worksheets("STATEMENT").Range(cel)
        .NumberFormat = "0.0000"
        .Value = .Value
    End With
Next

'----------------------------------------------------DOCUMENT SHEET
CellArray = Array("I:I", "BI:BJ")

For Each cel In CellArray
    With Worksheets("DOCUMENT").Range(cel)
        .NumberFormat = "0"
        .Value = .Value
    End With
Next

CellArray = Array("R:AC", "AJ:BC", "BH:BH")

For Each cel In CellArray
    With Worksheets("DOCUMENT").Range(cel)
        .NumberFormat = "0.00"
        .Value = .Value
    End With
Next

'----------------------------------------------------EX-ANTE REPORT SHEET
CellArray = Array("D:E", "AJ:AJ", "AN:AO")

For Each cel In CellArray
    With Worksheets("EX-Ante Report").Range(cel)
        .NumberFormat = "0"
        .Value = .Value
    End With
Next

CellArray = Array("O:P", "T:T", "Z:AB", "AD:AI")

For Each cel In CellArray
    With Worksheets("EX-Ante Report").Range(cel)
        .NumberFormat = "0.00"
        .Value = .Value
    End With
Next

ActiveWorkbook.Save

'----------------------------------------------------EX-ANTE CONSUMPTION SHEET
CellArray = Array("F:F")

For Each cel In CellArray
    With Worksheets("EX-Ante Consumption").Range(cel)
        .NumberFormat = "0%"
        .Value = .Value
    End With
Next

ActiveWorkbook.Save

End Sub




