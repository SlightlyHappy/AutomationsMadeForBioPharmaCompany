Attribute VB_Name = "Module41"
Sub ImportAndInsertMatchProxyCities()

    Dim sourceWorkbook As Workbook
    Dim sourceSheet As Worksheet
    Dim targetWorkbook As Workbook
    Dim sheetName As String
    Dim sourcePath As String

    ' --- Fixed File and Sheet Names ---
    sheetName = "Proxy cities"
    sourcePath = ThisWorkbook.Path & "\All Cost Estimate Line Item Help Text - Final Version.xlsx"
    ' --- End Fixed Names ---

    ' Import "Proxy cities"
    Set sourceWorkbook = Workbooks.Open(sourcePath)
    Set sourceSheet = sourceWorkbook.Worksheets(sheetName)
    Set targetWorkbook = ThisWorkbook
    Set targetSheet = targetWorkbook.Sheets.Add(After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count))
    targetSheet.Name = sheetName
    sourceSheet.UsedRange.Copy Destination:=targetSheet.Range("A1")
    sourceWorkbook.Close False

    ' --- Data Matching and Import (Columns B & C) ---

    With targetWorkbook.Worksheets("TransposedValues")
        .Columns("D:D").Insert Shift:=xlToRight ' Insert two new columns
        .Columns("D:D").Insert Shift:=xlToRight
        .Cells(1, "D").Value = "Proxy Home Country"
        .Cells(1, "E").Value = "Proxy Home City"
        
        lastRowProxy = targetWorkbook.Worksheets(sheetName).UsedRange.Rows.Count
        lastRowExtract = .UsedRange.Rows.Count

        For i = 2 To lastRowExtract
            For j = 2 To lastRowProxy
                If .Cells(i, "B").Value = targetWorkbook.Worksheets(sheetName).Cells(j, "A").Value And _
                   .Cells(i, "C").Value = targetWorkbook.Worksheets(sheetName).Cells(j, "B").Value Then

                    .Cells(i, "D").Value = targetWorkbook.Worksheets(sheetName).Cells(j, "C").Value
                    .Cells(i, "E").Value = targetWorkbook.Worksheets(sheetName).Cells(j, "D").Value
                    Exit For
                End If
            Next j
        Next i

    ' --- Data Matching and Import (Columns F & G) ---
        .Columns("H:H").Insert Shift:=xlToRight ' Insert two new columns
        .Columns("H:H").Insert Shift:=xlToRight
        .Cells(1, "H").Value = "Proxy Host Country"
        .Cells(1, "I").Value = "Proxy Host City"

        For i = 2 To lastRowExtract
            For j = 2 To lastRowProxy
                If .Cells(i, "F").Value = targetWorkbook.Worksheets(sheetName).Cells(j, "A").Value And _
                   .Cells(i, "G").Value = targetWorkbook.Worksheets(sheetName).Cells(j, "B").Value Then

                    .Cells(i, "H").Value = targetWorkbook.Worksheets(sheetName).Cells(j, "C").Value
                    .Cells(i, "I").Value = targetWorkbook.Worksheets(sheetName).Cells(j, "D").Value
                    Exit For
                End If
            Next j
        Next i
    End With

    MsgBox "Matching and import complete!"

End Sub


