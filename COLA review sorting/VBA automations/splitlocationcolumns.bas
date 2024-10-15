Attribute VB_Name = "Module31"
Sub SplitAndExtractData()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim splitValues() As String

    ' Define the worksheet to work with
    Set ws = ThisWorkbook.Sheets("TransposedValues")

    ' Find the last row with data in columns B and D
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    lastRowD = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row

    ' Process column B (including the first row)
    ws.Columns("C").Insert Shift:=xlToRight ' Insert new column

    For i = 1 To lastRow  ' Start from row 1 (the first row with data)
        splitValues = Split(ws.Cells(i, "B").Value, "/")
        If UBound(splitValues) >= 1 Then
            ws.Cells(i, "C").Value = Trim(splitValues(1)) ' Extract the second value
            ws.Cells(i, "B").Value = Trim(splitValues(0)) ' Overwrite with the first value
        End If
    Next i

    ' Process column D (including the first row)
    ws.Columns("E").Insert Shift:=xlToRight ' Insert new column

    For i = 1 To lastRowD ' Start from row 1 (the first row with data)
        splitValues = Split(ws.Cells(i, "D").Value, "/")
        If UBound(splitValues) >= 1 Then
            ws.Cells(i, "E").Value = Trim(splitValues(1)) ' Extract the second value
            ws.Cells(i, "D").Value = Trim(splitValues(0)) ' Overwrite with the first value
        End If
    Next i

End Sub

