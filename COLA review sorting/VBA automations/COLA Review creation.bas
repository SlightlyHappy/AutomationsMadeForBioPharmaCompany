Attribute VB_Name = "Module6"
Sub CopyColumnsToCOLAReview()
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim lastRow As Long

    ' Set references to sheets
    Set sourceSheet = ThisWorkbook.Sheets("ACS Extract")
    On Error Resume Next ' Handle if the sheet doesn't exist
    Set targetSheet = ThisWorkbook.Sheets("COLA Review")
    On Error GoTo 0

    If targetSheet Is Nothing Then
        ' Create the COLA Review sheet if it doesn't exist
        Set targetSheet = ThisWorkbook.Sheets.Add(After:=sourceSheet)
        targetSheet.Name = "COLA Review"
    End If
    
    ' Find the last row with data in the source sheet
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "A").End(xlUp).Row

    ' Copy columns A to K directly
    sourceSheet.Range("A1:K" & lastRow).Copy Destination:=targetSheet.Range("A1")

    ' Copy columns P, R, and X individually
    sourceSheet.Range("P1:P" & lastRow).Copy Destination:=targetSheet.Range("L1")
    sourceSheet.Range("R1:R" & lastRow).Copy Destination:=targetSheet.Range("M1")
    sourceSheet.Range("X1:X" & lastRow).Copy Destination:=targetSheet.Range("N1")
End Sub

