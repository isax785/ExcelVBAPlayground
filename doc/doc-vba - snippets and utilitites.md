# VBA - Snippets and Utilities

- [VBA - Snippets and Utilities](#vba---snippets-and-utilities)
  - [Screen Updating](#screen-updating)
  - [Cursor](#cursor)
  - [Progress Bar](#progress-bar)
  - [Sub](#sub)
  - [Interrupt a Sub](#interrupt-a-sub)
  - [Selection position](#selection-position)
  - [Concatenate Range of Cells](#concatenate-range-of-cells)
  - [Hexadecimal to VBA Color](#hexadecimal-to-vba-color)
  - [Goal Seek](#goal-seek)
  - [Copy-Paste Results on Table](#copy-paste-results-on-table)
  - [Select Range and Cycle on It](#select-range-and-cycle-on-it)
  - [Delete Rows](#delete-rows)
  - [Progress Bar](#progress-bar-1)
  - [Select Cells and Highlight Cells on Another Column](#select-cells-and-highlight-cells-on-another-column)

---


## Screen Updating

To avoid the **screen** changing during the simulation, then reducing the computation time, at the beginning of the macro:

> Application.ScreenUpdating = False

## Cursor

To hidden the cursor (i.e. convert the arrow to the loading symbol), write at the beginning of the macro:

> Application.Cursor = xlWait 'cursor is hidden

At the end of the macro, **the cursor must be restored**:

> Application.Cursor = xlDefault 'cursor is shown

## Progress Bar

To write a message on the status bar at the bottom of the sheet:

> Application.StatusBar = "This is the message"

Such message can be updated in any part of the macro (e.g. inside a _for_ cycle): then the percentage of the progress can be calculated and shown. Below an example:

```vb
Dim rng As Range
Dim total, progress, perc As Integer: progress = 0

Range(Range([cellID]), Range([cellID]).End(xlDown)).Select
Set rng = Selection
total = rng.Count

For Each cell In rng.Cells
	progress = progress + 1
	perc = progress / total * 100
	Application.StatusBar = perc & "% completed"
	Next cell
```

## Sub

```vb
Sub nested_sub()
	debug.print "does nothing"
End Sub

Sub main()
    Call nested_sub
End Sub
    
```

## Interrupt a Sub

Below the code:

```vb
Sub interrupt(sheetName As String)
    ' Checks to avoid using the macro in the wrong sheet
    If Not ActiveSheet.Name = sheetName Then
        MsgBox "You're in the wrong sheet!" & vbNewLine & "Please switch to the sheet: " & sheetName, vbCritical
        'Exit Sub
        End
    End If
End Sub
```

This can be called inside another _Sub_.

## Selection position

At the end of a macro, it could be useful to place the cell selection in a specific cell in order to avoid to move manually when the simulation ends. For example:

> Range("A1").Select   'selection of the worksheet top-left cell

The same can be done to move to a specific sheet:

> Sheets(sheetName).Activate
>
> Range("A1").Select

## Concatenate Range of Cells

```vb
Function CONCATENATE_INTERVAL(interval As Range, Optional splitter As String) As String
    Dim cnt As Long
    Dim cell as Range
        
    Application.Volatile
    
    cnt = 0
    For Each cell In interval.Cells
        If IsEmpty(cell) = False Then
            cnt = cnt + 1
            If cnt = 1 Then
                CONCATENATE_INTERVAL = cell.Value
            Else
                CONCATENATE_INTERVAL = CONCATENATE_INTERVAL & splitter & cell.Value
            End If
        End If
    Next cell
  
End Function
```

> <u>**Example**</u>
>
> ESK_CONCATENA_INTERVALLO(A1:A100; " ;")

## Hexadecimal to VBA Color

```vb
Function hexa_color(ByVal hexa) 'Returns -1 in case of error

    'Convert a hexadecimal color to a Color value - Excel-Pratique.com
    'www.excel-pratique.com/en/vba_tricks/hexadecimal-color-function

    If Len(hexa) = 7 Then hexa = Mid(hexa, 2, 6) 'If color with #
    
    If Len(hexa) = 6 Then

        num_array = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f")
        
        char1 = LCase(Mid(hexa, 1, 1))
        char2 = LCase(Mid(hexa, 2, 1))
        char3 = LCase(Mid(hexa, 3, 1))
        char4 = LCase(Mid(hexa, 4, 1))
        char5 = LCase(Mid(hexa, 5, 1))
        char6 = LCase(Mid(hexa, 6, 1))
        
        For i = 0 To 15
            If (char1 = num_array(i)) Then position1 = i
            If (char2 = num_array(i)) Then position2 = i
            If (char3 = num_array(i)) Then position3 = i
            If (char4 = num_array(i)) Then position4 = i
            If (char5 = num_array(i)) Then position5 = i
            If (char6 = num_array(i)) Then position6 = i
        Next
        
        If IsEmpty(position1) Or IsEmpty(position2) Or IsEmpty(position3) Or IsEmpty(position4) Or IsEmpty(position5) Or IsEmpty(position6) Then
            hexa_color = -1
        Else
            hexa_color = RGB(position1 * 16 + position2, position3 * 16 + position4, position5 * 16 + position6)
        End If
        
    Else
        hexa_color = -1
    End If
    
End Function
```

## Goal Seek

> Range([main]).GoalSeek goal:=Range([goal]), ChangingCell:=Range([variation])

Where:

- _[main]_ - cell that must reach the goal value
- _[goal]_ - cell containing the goal value
- _[variation]_ - cell to be varied to make _[main]_ reaching _[goal]_

To increase the chance to reach the convergence and to reduce the computational time, the _[main]_ value initialization is highly recommended:

> Range([main]) = [initialValue]

## Copy-Paste Results on Table

```vb
Sub copyPasteSim()
    ' 1. copy t0 and t1 of a previously simulated scenario
    ' 2. paste to input for a further simulation
    ' 3. copy the results
    ' 4. paste the results to overwrite the previou ones
    
    ' NOTE: any values can be added to the results and re-simulated with this macro
    
    Dim c As Range
    
    ' copy input values
    Set c = Selection
    
    'Range("BM15:BM16").Select
    Range(c, c.Offset(1, 0)).Select
    Selection.Copy
    
    ' paste input values to simulation input
    Range("AZ3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False  'paste values only
            
    ' copy simulation results
    Range("BF15").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    
    'Range("BM15").Select
    c.Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Application.CutCopyMode = False    'removes the highlight of the copied area
    c.Select
        
End Sub

```

## Select Range and Cycle on It

```vb
Dim rng As Range
	Range(Range([cellID]), Range([cellID]).End(xlDown)).Select
Set rng = Selection

For Each cell In rng.Cells
	debug.print cell.value
Next cell
```

## Delete Rows

```vb
Sub delete_rows()
  Dim s As String
  Dim i As Integer
    Dim i_max as Integer: i_max = 2500
    
  i = Selection.row
  Do While i < i_max
      s = Str(i)
      Range(i & ":" & i).Select
      Selection.Delete Shift:=xlUp
      i = i + 1
  Loop
  End Sub
```

## Progress Bar

At first, the total number of simulations to be carried out is computed. Then, by cycling on that number, the progress percentage is shown.

```vb
dim runs as Integer 'total number of simualations
If IsEmpty(Range("A1")) = False And IsEmpty(Range("A1").Offset(1, 0)) = False Then
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    runs = Selection.Count
End If

i = 0

Do While i < runs do
    '[do something]
    i = i + 1   
    Application.StatusBar = Round(i / runs * 100, 2) & "% completed"
Loop
```

## Select Cells and Highlight Cells on Another Column

[Source]](https://danwagner.co/how-to-select-a-cell-in-one-column-and-highlight-the-corresponding-cell-in-another-column/)

![this-event-code-goes-in-the-worksheet-not-a-regular-module](this-event-code-goes-in-the-worksheet-not-a-regular-module.png)

```vb
Option Explicit
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Debug.Print "here!"
    Dim wksLookups As Worksheet
    Dim lngFirstRow As Long, lngLastRow As Long
    Dim col_nr As Integer: col_nr = 1
    
    'Get the Worksheet so we can confidently identify
    'the Range that we'll be highlighting (if need be)
    Set wksLookups = Target.Parent
    
    'Clear the background color in column col_nr
    wksLookups.Columns(col_nr).Interior.ColorIndex = xlNone
    
    'First, we need to check that the selection
    'occurred in column col_nr AND was in a row after
    'row 1
    If Target.Column <> col_nr And Target.Row > 1 Then
        
        'Find the first and last rows of the selected
        'range in column A
        lngFirstRow = Target.Row
        lngLastRow = Target.Rows.Count + (Target.Row - 1)
        
        'Highlight the corresponding cell(s) in column col_nr
        With wksLookups
            .Range(.Cells(lngFirstRow, col_nr), _
                   .Cells(lngLastRow, col_nr)).Interior.Color = RGB(35, 255, 110)
        End With
        
    End If
    
End Sub


```


