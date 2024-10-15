Attribute VB_Name = "Module5"
Sub MoveValues11()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cell As Range
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim textInParentheses As String
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("ACS Extract")
    
    ' Find the last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Create a Regular Expression object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    
    ' Loop through each cell in column A
    For i = 1 To lastRow
        ' Check if the cell contains "(+)"
        If InStr(1, ws.Cells(i, 1).Value, "(+)") > 0 Then
            ' Use regular expression to find the value between parentheses from the right side of the cell
            regex.Pattern = "\(([^)]+)\)$"
            Set matches = regex.Execute(ws.Cells(i, 1).Value)
            ' If there's a match
            If matches.Count > 0 Then
                ' Store the text within parentheses
                textInParentheses = Trim(matches(0).SubMatches(0))
                ' Append the value to the cell in front of it
                ws.Cells(i, 1).Value = Replace(ws.Cells(i, 1).Value, matches(0), "")
                ws.Cells(i, 1).Value = Trim(ws.Cells(i, 1).Value)
                ' Retain the original value in column B and append the extracted value with a comma
                ws.Cells(i, 2).Value = Trim(ws.Cells(i, 2).Value) & ", " & Trim(textInParentheses)
            End If
        End If
    Next i
    
    ' After the initial loop, replace specific values in column A
    For i = 1 To lastRow
        If ws.Cells(i, 1).Value = "Designated Home Country" Then
            ws.Cells(i, 1).Value = "Home Country / Home City"
        ElseIf ws.Cells(i, 1).Value = "Family Status (At Home / At Post)" Then
            ws.Cells(i, 1).Value = "Family Status (Home Country / Host Country)"
        End If
    Next i
    
    ' Cleanup
    Set regex = Nothing
    Set matches = Nothing
End Sub

