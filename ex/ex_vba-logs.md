# Logs

- [Logs](#logs)
  - [Static Log Array](#static-log-array)
  - [Dynamic Log Array](#dynamic-log-array)
  - [Logging with a Collection](#logging-with-a-collection)

---

## Static Log Array

```vb
Sub WriteLogs()
    Dim logs() As String
    Dim i As Long
    Dim logCount As Long
    Dim startCell As Range
    
    'Set the cell where logs will start to be written
    Set startCell = ThisWorkbook.Worksheets("Sheet1").Range("A1")
    
    logCount = 10 'example loop count, replace as needed
    
    'Initialize the array to the number of logs
    ReDim logs(1 To logCount)
    
    'Collect logs inside a loop
    For i = 1 To logCount
        logs(i) = "Log entry number " & i
    Next i
    
    'Write logs to the sheet starting from the startCell downward
    startCell.Resize(logCount, 1).Value = Application.Transpose(logs)
    
End Sub
```

## Dynamic Log Array

```vb
Sub WriteLogs_UnknownCount()

    Dim logs() As String
    Dim logIndex As Long
    Dim i As Long
    Dim startCell As Range
    
    Set startCell = ThisWorkbook.Worksheets("Sheet1").Range("A1")
    
    'Start with an uninitialized dynamic array
    ReDim logs(1 To 1)
    logIndex = 1
    
    '----------------------------------------------
    ' Example: real loop of unknown variable length
    ' Assume we loop through a column until an empty cell is found
    '----------------------------------------------
    i = 1
    Do While ThisWorkbook.Worksheets("Sheet1").Range("C" & i).Value <> ""
        
        'Store log
        logs(logIndex) = "Processed value: " & _
                         ThisWorkbook.Worksheets("Sheet1").Range("C" & i).Value
        
        'Increase array size for next element
        logIndex = logIndex + 1
        ReDim Preserve logs(1 To logIndex)
        
        i = i + 1
    Loop
    
    'Remove last empty element created by the final ReDim
    ReDim Preserve logs(1 To logIndex - 1)
    
    '----------------------------------------------
    ' WRITE LOGS TO WORKSHEET
    '----------------------------------------------
    startCell.Resize(UBound(logs), 1).Value = Application.Transpose(logs)

End Sub
```

## Logging with a Collection

```vb
Sub WriteLogs_Collection()

    Dim logs As New Collection
    Dim i As Long
    Dim startCell As Range
    Dim ws As Worksheet
    Dim logItem As Variant
    Dim outputArr() As String
    Dim idx As Long
    
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    Set startCell = ws.Range("A1")
    
    '---------------------------------------------------------
    ' Example loop: runs until an empty cell in Column C is found
    '---------------------------------------------------------
    i = 1
    Do While ws.Range("C" & i).Value <> ""
        
        logs.Add "Processed value: " & ws.Range("C" & i).Value
        
        i = i + 1
    Loop
    
    '---------------------------------------------------------
    ' Transfer collection to an array for fast writing to sheet
    '---------------------------------------------------------
    If logs.Count = 0 Then Exit Sub ' nothing to write
    
    ReDim outputArr(1 To logs.Count)
    
    idx = 1
    For Each logItem In logs
        outputArr(idx) = logItem
        idx = idx + 1
    Next logItem

    '---------------------------------------------------------
    ' Write logs to worksheet (efficient single write)
    '---------------------------------------------------------
    startCell.Resize(logs.Count, 1).Value = Application.Transpose(outputArr)

End Sub
```


---

[EX MOC](./ex-00_MOC.md)
