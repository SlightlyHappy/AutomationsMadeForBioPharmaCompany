Attribute VB_Name = "Module1"
Sub ImportBASFilesAndProcessSheets()
    Dim fd As FileDialog
    Dim folderPath As String
    Dim fileName As String
    Dim vbProj As Object
    Dim vbComp As Object
    Dim fullPath As String
    Dim ws As Worksheet

    ' Initialize the FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    With fd
        .Title = "Select Folder Containing .bas Files"
        .AllowMultiSelect = False
        If .Show = -1 Then
            folderPath = .SelectedItems(1) & "\"
        Else
            MsgBox "No folder selected.", vbExclamation
            Exit Sub
        End If
    End With

    ' Reference the VBProject dynamically
    Set vbProj = ThisWorkbook.VBProject

    ' Loop through each .bas file in the folder and import it
    fileName = Dir(folderPath & "*.bas")
    Do While fileName <> ""
        fullPath = folderPath & fileName
        ' Dynamically import the .bas file
        Set vbComp = vbProj.VBComponents.Import(fullPath)
        fileName = Dir
    Loop

    ' Run the Sub ExtractData11
    Application.Run "ExtractData11"

    ' Run the Sub MoveValues11
    Application.Run "MoveValues11"

    ' Run the Sub TransposeData
    Application.Run "TransposeData"

    ' Run the Sub SplitAndExtractData
    Application.Run "SplitAndExtractData"

    ' Run the Sub ImportAndInsertMatchProxyCities
    Application.Run "ImportAndInsertMatchProxyCities"
    
    ' Run the Sub ImportAPROMonthlyAndLookup
    Application.Run "ImportAPROMonthlyAndLookup"
    
    ' Delete all sheets except specified ones
    Application.DisplayAlerts = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "APROreport" And ws.Name <> "Proxy cities" And ws.Name <> "ACS Extract" And ws.Name <> "TransposedValues" Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True

    MsgBox "All tasks completed successfully.", vbInformation
End Sub

