Attribute VB_Name = "Module51"
Sub ImportAPROMonthlyAndLookup()

    ' Get directory and import APROreport (same as before)
    Dim CurrentWorkbookPath As String
    CurrentWorkbookPath = ThisWorkbook.Path
    Dim APROmonthlyPath As String
    APROmonthlyPath = CurrentWorkbookPath & "\APROmonthly.xlsx"

    Dim APROWorkbook As Workbook
    Set APROWorkbook = Workbooks.Open(APROmonthlyPath, ReadOnly:=True)
    APROWorkbook.Worksheets("Sheet1").Copy Before:=ThisWorkbook.Sheets(1)
    APROWorkbook.Close SaveChanges:=False
    ThisWorkbook.Worksheets(1).Name = "APROreport"

    ' Sheets and ranges for lookup
    Dim TransposedValuesSheet As Worksheet ' Changed name here
    Dim APROreportSheet As Worksheet
    Set TransposedValuesSheet = ThisWorkbook.Sheets("TransposedValues") ' Changed name here
    Set APROreportSheet = ThisWorkbook.Sheets("APROreport")

    Dim TransposedValuesRange As Range ' Changed name here
    Dim APROreportRange As Range
    Set TransposedValuesRange = TransposedValuesSheet.Range("A2", TransposedValuesSheet.Cells(Rows.Count, "A").End(xlUp)) ' Changed name here
    Set APROreportRange = APROreportSheet.Range("B2", APROreportSheet.Cells(Rows.Count, "B").End(xlUp))

    ' Insert new column in TransposedValues ' Changed name here
    TransposedValuesSheet.Columns(1).Insert Shift:=xlToRight ' Changed name here
    TransposedValuesSheet.Cells(1, 1).Value = "Workday ID" ' Rename the new column

    ' Loop through TransposedValues values and perform lookup ' Changed name here
    Dim TransposedValuesCell As Range ' Changed name here
    For Each TransposedValuesCell In TransposedValuesRange ' Changed name here
        Dim lookupValue As String
        lookupValue = Split(TransposedValuesCell.Value, ",")(0) ' Get text before ","

        Dim foundCell As Range
        Set foundCell = APROreportRange.Find(lookupValue, LookIn:=xlValues, LookAt:=xlWhole)
        If Not foundCell Is Nothing Then
            TransposedValuesCell.Offset(0, -1).Value = foundCell.Offset(0, -1).Value ' Copy from APROreport
        End If
    Next TransposedValuesCell ' Changed name here

End Sub

