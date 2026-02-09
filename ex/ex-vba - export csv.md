# Export `csv`

A VBA macro that exports all sheets of an Excel file to separate CSV files, with a user-defined separator. The macro will prompt the user to input the desired separator and will then save each sheet as a CSV file with the given separator.

```vb
Sub ExportAllSheetsToCSV()
    Dim ws As Worksheet
    Dim csvFileName As String
    Dim separator As String
    Dim folderPath As String
    Dim currentSheet As Worksheet
    Dim fileSystemObject As Object
    Dim textStream As Object
    Dim cell As Range
    Dim rowContent As String
    
    ' Prompt user for the separator
    ' separator = InputBox("Enter the separator for CSV (e.g., comma, semicolon, etc.):", "CSV Separator", ",")
    
    separator = ";"

    ' Validate the separator
    If separator = "" Then
        MsgBox "No separator entered. Operation cancelled.", vbExclamation
        Exit Sub
    End If

    ' Prompt user to select the folder to save CSV files
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder to Save CSV Files"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With

    ' Loop through each sheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        csvFileName = folderPath & "\" & ws.Name & ".csv"
        Debug.Print "Exporting" & csvFileName ''''
        
        ' Create FileSystemObject and TextStream
        Set fileSystemObject = CreateObject("Scripting.FileSystemObject")
        Set textStream = fileSystemObject.CreateTextFile(csvFileName, True, False)

        ' Loop through each row and column in the sheet
        For Each cell In ws.UsedRange.Cells
            If cell.Column = 1 Then
                ' Debug.Print cell.Text '''''
                If IsEmpty(cell) Then
                    Exit For
                End If
                
                
                If cell.Row > 1 Then textStream.Write vbCrLf
                rowContent = ""
            Else
                rowContent = rowContent & separator
            End If
            rowContent = rowContent & cell.Text
            If cell.Column = ws.UsedRange.Columns.Count Then textStream.Write rowContent
        Next cell
        
        Debug.Print "Done!!"
        ' Close the text stream
        textStream.Close
    Next ws

    ' Clean up
    Set fileSystemObject = Nothing
    Set textStream = Nothing

    MsgBox "All sheets have been exported as CSV files to " & folderPath, vbInformation
End Sub
```