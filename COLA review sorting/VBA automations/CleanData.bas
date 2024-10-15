Attribute VB_Name = "Module2"
Sub ReorganizeCompensationTable()
    Dim ws As Worksheet

    ' Assuming you want to modify the active sheet
    Set ws = ActiveSheet

    ' 1. Combine split text in row 27
    ws.Range("A27").Value = ws.Range("A27").Value & " " & ws.Range("C27").Value & " " & ws.Range("D27").Value
    ws.Range("C27:D27").ClearContents ' Remove original split text

    ' 2. Shift columns D to B (Rows 4-23)
    ws.Range("D4:D23").Cut ws.Range("B4")

    ' 3. Merge and move D25 & E25 to B25
    With ws.Range("B25")
        .Value = ws.Range("D25").Value & " " & ws.Range("E25").Value ' Combine the values with a space
        .Merge
    End With
    ws.Range("D25:E25").ClearContents

    ' 4. Move F27 to B27
    ws.Range("F27").Cut ws.Range("B27")
    
    ' 5. Exception Handling for A20 (Final Version)
    indexStart = InStr(ws.Range("A20").Value, "(Index = ") ' Find start of index text
    If indexStart > 0 Then
        indexEnd = InStr(indexStart, ws.Range("A20").Value, ")") ' Find end of index text
        If indexEnd > 0 Then
            ws.Range("B20").Value = ws.Range("B20").Value & " " & Mid(ws.Range("A20").Value, indexStart, indexEnd - indexStart + 1)
            ws.Range("A20").Value = Left(ws.Range("A20").Value, indexStart - 1) ' Remove index text and the extra space
        End If
    End If

End Sub
