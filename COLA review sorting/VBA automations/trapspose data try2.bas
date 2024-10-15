Attribute VB_Name = "Module3"
Sub TransposeData()
    Dim wsSource As Worksheet
    Dim wsDestination As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim stopTransposing As Boolean
    
    ' Set the source worksheet
    Set wsSource = ThisWorkbook.Sheets("ACS Extract") ' Change "Sheet1" to your source sheet name
    
    ' Add a new sheet and rename it to "TransposedValues"
    Set wsDestination = ThisWorkbook.Sheets.Add
    wsDestination.Name = "TransposedValues"
    
    ' Find the last row in the source sheet
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    
    ' Loop through each row in the source sheet
    For i = 1 To lastRow
        ' If cell in column A is not empty and it matches the first header, copy the header to destination sheet
        If wsSource.Cells(i, 1).Value <> "" Then
            If Not stopTransposing Then
                wsDestination.Cells(1, j + 1).Value = wsSource.Cells(i, 1).Value
                j = j + 1
                If wsSource.Cells(i, 1).Value = "(+) Cost of Living Allowance" Then
                    stopTransposing = True
                End If
            End If
        End If
    Next i
    
    ' Loop through each row in the source sheet starting from row 2
    For i = 2 To lastRow
        ' If cell in column B is not empty, transpose corresponding values to destination sheet
        If wsSource.Cells(i, 2).Value <> "" Then
            For j = 1 To wsDestination.Cells(1, wsDestination.Columns.Count).End(xlToLeft).Column
                If wsDestination.Cells(1, j).Value <> "" Then
                    If wsDestination.Cells(1, j).Value = wsSource.Cells(i, 1).Value Then
                        wsDestination.Cells(wsDestination.Cells(wsDestination.Rows.Count, j).End(xlUp).Row + 1, j).Value = wsSource.Cells(i, 2).Value
                    End If
                End If
            Next j
        End If
    Next i
End Sub

