Attribute VB_Name = "Module4"
Sub ExtractData11()
    Dim ws As Worksheet
    Dim searchValues As Variant
    Dim searchRange As Range
    Dim foundCell As Range
    Dim lastRow As Long
    Dim extractSheet As Worksheet
    Dim newRow As Long
    
    ' Define the search values
    searchValues = Array("Name / First name", "Home Country / Home City", "Host Country / Host City", _
                         "Family Status (Home Country / Host Country)", "Family Status (At Home / At Post)", _
                         "Currency", "Annual Gross Base Salary", "Cost of living Allowance", _
                         "Designated Home Country")
    
    ' Create a new sheet for the extracted data
    Set extractSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    extractSheet.Name = "ACS Extract"
    newRow = 1
    
    ' Loop through each sheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> extractSheet.Name Then ' Exclude the extract sheet itself
            ' Find the last row in column A of the current sheet
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            
            ' Set the search range to column A of the current sheet
            Set searchRange = ws.Range("A1:A" & lastRow)
            
            ' Loop through each search value
            For Each Value In searchValues
                ' Look for the search value in column A (fuzzy match)
                Set foundCell = searchRange.Find(What:=Value, LookIn:=xlValues, LookAt:=xlPart)
                
                If Not foundCell Is Nothing Then
                    ' If the search value is found, copy the value from column A and B to the extract sheet
                    extractSheet.Cells(newRow, 1).Value = foundCell.Value
                    ' Check if the value is "Family Status (At Home / At Post)" or "Designated Home Country"
                    If Value = "Family Status (At Home / At Post)" Or Value = "Designated Home Country" Then
                        ' Copy the value from column B and C if found in the same row
                        extractSheet.Cells(newRow, 2).Value = foundCell.Offset(0, 1).Value
                        extractSheet.Cells(newRow, 3).Value = foundCell.Offset(0, 2).Value
                    Else
                        ' Otherwise, copy the value from column B
                        extractSheet.Cells(newRow, 2).Value = foundCell.Offset(0, 1).Value
                    End If
                    newRow = newRow + 1
                End If
            Next Value
        End If
    Next ws
    
    MsgBox "Data extraction completed.", vbInformation
End Sub


