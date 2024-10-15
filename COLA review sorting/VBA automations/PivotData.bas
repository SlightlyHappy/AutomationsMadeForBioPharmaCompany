Attribute VB_Name = "Module3"
Sub PivotTableData()
    Dim lastRow As Long
    Dim cycleLength As Long
    Dim i As Long, j As Long
    Dim newRow As Long
    Dim headerRng As Range
    Dim wsSource As Worksheet, wsPivot As Worksheet
    Dim dictHeaders As Object

    Set wsSource = ThisWorkbook.Worksheets("ACS Extract")
    Set wsPivot = ThisWorkbook.Worksheets.Add(After:=wsSource)
    wsPivot.Name = "PivotedData"

    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row

    Set dictHeaders = CreateObject("Scripting.Dictionary")

    ' Populate dictionary with unique headers and their column positions
    For i = 1 To lastRow
        headerName = wsSource.Cells(i, 1).Value
        If Not dictHeaders.Exists(headerName) Then
            dictHeaders.Add headerName, dictHeaders.Count + 1
            cycleLength = cycleLength + 1 ' Increment cycle length with each new unique header
        End If
    Next i

    ' Write headers to new sheet
    For Each headerName In dictHeaders.Keys()
        wsPivot.Cells(1, dictHeaders(headerName)).Value = headerName
    Next headerName

    ' Pivot the data, creating new rows for each cycle
    newRow = 2
    For i = 1 To lastRow Step cycleLength
        For j = 0 To cycleLength - 1
            wsPivot.Cells(newRow, j + 1).Value = wsSource.Cells(i + j, 2).Value
        Next j
        newRow = newRow + 1
    Next i
End Sub

