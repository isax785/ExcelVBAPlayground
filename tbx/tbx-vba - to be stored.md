# VBA TBX to Be Stored

- [VBA TBX to Be Stored](#vba-tbx-to-be-stored)
- [Undefined Actions](#undefined-actions)
- [Notes](#notes)
- [Tables](#tables)
- [Snippets](#snippets)
  - [`With` Loop](#with-loop)
  - [Copy-Paste Range](#copy-paste-range)
  - [Select Case](#select-case)
  - [Text File](#text-file)
  - [Dialog for Folder Selection](#dialog-for-folder-selection)
  - [Conditional Formatting](#conditional-formatting)
    - [Cell Color](#cell-color)
  - [Charts](#charts)
    - [Copy-Paste Chart](#copy-paste-chart)
    - [Set Chart](#set-chart)

---

> From the book `100 Excel VBA Simulations`

# Undefined Actions

- `rng.FormulaArray`

# Notes

- rows and columns counter starts from `1`

# Tables

| **General**                            |                                        |
| ---                                    | ---                                    |
| Multiple statements in a single row    | *`[statement] : [statement]`*          |
| Label placed anywhere within the macro | *`[label]:`*                           |
| Jump to a specific label `[label]`     | *`GoTo [label]`*                       |
| Round to the lower integer             | `i = Int([value])`                     |
| Random integer number between 0 and N  | `Int(Rnd * (N+1))`                     |
| Module -> `int`                        | *`[int] Mod [int]`*                    |


| **Keywords**                           |                                        |
| ---                                    | ---                                    |
| Random value in the `[0,1]` range, `0` and `1` included     | `Rnd`             |
| Timer counting **seconds** since the midnight of the machine | `Timer`          |
| Newline keyword for string             | `vbcr`                                 |
| Tab keyword for string                 | `vbTab`                                |
| Prevention of system blocking (inside a loop) | `DoEvents`                      |
| Break loop                             | *`Exit [loop]`*                        |
| Row code separator                     | `:`                                    |
| Make global variable declaration (at the top of the module) | `Option Explicit` |
| Serial number of current date and time (`2/26/2026 3:23:30 PM`) | `Now`         |
| Calculate all open workbooks           | `Calculate`                            |
| All the spreadsheet cells              | `Cells`                                |
| Currently active sheet                 | `ActiveSheet`                          |


| **Application**                        |                                       |
| ---                                    | ---                                   |
| Calculate all open workbooks	         | `Application.Calculate` or `Calculate`| 
| Calculate a specific worksheet	     | `Worksheets(1).Calculate`             |
| Calculate a specified range	         | `Worksheets(1).Rows(2).Calculate`     |
| Freeze screen while executing the macro | `Application.ScreenUpdating = False` | 
| Unlock screen while executing the macro | `Application.ScreenUpdating = True`  | 
| Force calculation to "manual" call      | `Application.Calculation = xlCalculationManual `    |
| Restore calculation to automatic call   | `Application.Calculation = xlCalculationAutomatic ` | 
| Alerts display/hide           | `Application.DisplayAlerts = True` or `=False` |


| **Declarations**                       |                                       |
| ---                                    | ---                                   |
| Variant: can store anything            | *`Dim [name] as Variant`*             |
| Inline declarations, single type       | *`Dim [varname] as [vartype]`*        |
| Inline declaration and assignment      | *`Dim [varname] as [vartype] : [varname] = [value]`*   |
| Inline declarations, multiple types    | *`Dim [varname] as [vartype], [varname] as [vartype]`* |
| Declare and fill array    | *`Dim [arrname] as Variant : [arrname] = Array([val], [val], ...)`* |
| Declare array of integers, undefined size | `arr() as Integer`                 |
| Declare array of `N-M+1` integers      | `Dim arr(M to N) as Integer`          |
| Declare array of `N-M+1` anything      | `Dim arr(M to N) as Variant`          |
| Re-declare array to set size (5)       | `ReDim arr(4)`                        |
| Re-define array size (upper bound only) without changing the contained data | `ReDim Preserve arr(10)` |
| Array as variant, then fill it         | `Dim arr as Variant : arr = Array(1, 2, 3)` |
| Size of 1D array -> `int`              | `UBound(arr)`                         |
| Size of multidimensional array         | *`UBound([arr], [dim])`*              |


| **String Manipulation**                |                                       |
| ---                                    | ---                                   |
| String concatenation (spaces are not automatically inserted) | *`"[string]" & "[string]" * [int]`* |
| Convert value to string                | *`CStr([val])`*                       |
| String to upper/lower case             | *`UCase([string])`* / *`LCase([string])`* |
| Separator to compose strings on multiple lines | `_`                           |
| Character to integer                   | `CInt(...)`                           |


| **MessageBox**                         |                                       |
| ---                                    | ---                                   |
| Open messagebox                        | *`MsgBox("[message]", [button-set], [box-title])`* |
| Messagebox button set                  | `vbOkCancel`, `vbYesNoCancel`, `vbYesNo` |
| Buttons signals                        | `vbOK`, `vbCancel`, `vbYes`, `vbNo`   |
| Conditional | *`If MsgBox("[message]", [button-set]) = [signal] Then [action] `* |
| Get messagebox output | `Dim msg as Variant` `msg = MsgBox(...)`               |
| Oputput cases                          | `Yes`    -> `Case 6`                  |
|                                        | `No`     -> `Case 7`                  |
|                                        | `Cancel` -> `Case 2`                  |


| **InputBox**                           |                                       |
| ---                                    | ---                                   |
| Open input box                        | *`Dim v as [type] : v = InputBox("[message]", ,[default])`* |
| | *`Application.InputBox (Prompt, Title, Default, Left, Top, HelpFile, HelpContextID, Type)`* |



| **Cells and Ranges**                   |                                       |
| ---                                    | ---                                   |
| Offset (`[row]` and `[col]` are incremental values) | `Range(...).Offset([row], [col])`     |
| Get address -> `str`                   | *`[range].Address`*                   |
| Select region                          | `Range(...).CurrentRegion`            |
| Region row and column count -> `int`   | `.CurrentRegion.Rows.Count` `.CurrentRegion.Columns.Count`  |
| Clear the region                       | `.CurrentRegion.Delete`               | 
| Access to cell value with coordinates  | *`Cells([row], [col])`*               |
| Define a range with cells coordinates  | *`Range(Cells([row], [col]), Cells([row], [col]))`* |
| Named range                            | *`Range("[name]")`*                   |
| Write a formula                        | *`Range(...).Formula = "=[formula]"`* | 
| *absolute in the sheet*                | `R1C1`                                |
| *relative to active position*          | `R[-1]C[1]`  (1 row upper, 1 column righ) |
| Write a R1C1 formula, i.e. relative notation | *`Range(...).FormulaR1C1 = "=[formula]"`*  | 
| Clear range content                    | `.ClearContents`                      |
| Clear entire column                    | `Range("A1").EntireColumn.Clear`      |
| Clear entire column content            | `Range("A1").EntireColumn.ClearContents`|
| Columns range                          | `Columns("A:C")`                      |
| Sorting                                | *`Range("[range-to-sort]").Sort Range("[first-sort-field]") [order]`*   |
| Sorting (`xlAscending` by default)  | `Range("A4:B270").Sort Range("B4")`      | 
| Sorting descending order (??)          | `Range("D4:E270").Sort Range("E4"), xlDescending`   |
| Access to column within a range        | `With rng` `.Columns(1) ...`          |
| Access to row within a range           | `With rng` `.Rows(1) ...`             |
| Assign a region to a range             | `Set rng = rng0.CurrentRegion`        |
| Resize range (both parameters are optional) | *`rng.Resize([RowSize], ColumnSize)`* |


| **Cells and Ranges Formatting**        |                                          |
| ---                                    | ---                                      |
| Formatting variable                    | *`Dim [name] as FormatCondition`*        |
| Databar conditional                    | *`Dim [name] as Databar`*                |
| Cell coloring                          | *`.Cell.Interior.ColorIndex = [value]`*  |
| Range cells coloring                   | *`.Cells.Interior.ColorIndex = [value]`* |
| Thick border                           | `.borderAround, xlThick`                 |
| Horizontal alignment                   | `.Cells.HorizontalAlignment = xlCenter`  |
| Border styling                        | `.Cells.Borders.LineStyle = xlContinuous` |
| Autofit column width                   | `Cells.EntireColumnutoFit`             |
| Delete conditional formatting          | `Range(...).FormatConditions.Delete`     |
| Percent formatting                     | `rng = FormatPercent([val], [decimals])` |


| **Loops and Conditionals**             |                                        | 
| ---                                    | ---                                    |
| `For` loop (`Dim i as Integer`)        | `For i = 0 to 5 ... Next i`            |
| `For` loop with defined step        | `For number As Double = 0 To 2 Step 0.25` |
| `For` loop break                       | `Exit For`                             |
| `With` loop (`Dim r as Range : Set r = Range(...)`) | `With r ... .Cells(1, 1) = ... End With` |
| `If` condition                       | *`If [condition] Then ... End If`*       |
| `If ... Else` condition           | *`If [condition] Then ... Else ... End If`* |
| `If ... ElseIf ... Else` condition | *`If [condition] Then ... ElseIf [condition] Then ... Else ... End If`* |
| Inline conditional assignment          | *`i = IIF([condition], [true-return], [false-return])`*             |
| `Do` loop - a break condition is mandatory  | *`Do ... Loop`*                   |
| Break `Do` loop                        | `Exit Do`                              | 
| `Do` loop with breking condition  | *`Do Until [condition] ... Loop`*           |
|                                   | *`Do ... Loop Until [condition]`*           |
| Select cases | *`Select case [var] Case [val]: ... Case [val]: ... End Select`* |


| **Worksheet Handling**                 |                                       |
| ---                                    | ---                                   |
| Add worksheet after the active one | `Dim wks as Worksheet : Set wks = Worksheets.Add( , ActiveSheet)` |
| Assign active sheet           | `Dim wks as Worksheet : Set wks = ActiveSheet` |
| Copy the active sheet                  | `wks.Copy, Sheets(Sheets.Count)`      |
| Select sheet on number                 | `Sheet1.Select`                       |
|                                        | `Sheets(1).Select`                    |

| **Workbook Functions**                 |                                       |
| ---                                    | ---                                   |
| reference to the current workbook | `ThisWorkbook`                             | 
| iterate over all the worksheets | `For Each ws In ThisWorkbook.Worksheets ... Next ws` |

| **Worksheet Functions**                |                                       |
| ---                                    | ---                                   |
| Call function                          | *`WorksheetFunction.[functioname(arguments)]`* |
| Count blank cells -> `int`             | *`.CountBlank([range])`*              | 
| Integer random between two values      | *`.RandBetween([int], [int])`*        |
| Count cells matchin a defined condition | *`.CountIf([range], "[condition]")`* |
| Set zoom on active window              | `ActiveWindow.Zoom = 130`             |
| Goal seek            | *`Range(...).GoalSeek [goal-value], [cell-to-change]`*  |
| Calculate average value                | *`.Average([range])`*                 |
| Calculate percentile vlue              | `.Percentile([range], [val])`         |
| Vertical lookup  | *`.VLookup([lookup-value], [table-array], [col-index-num], [range-lookup])`* |
|  |  |
|  |  |
|  |  |
|  |  |

| **Charts**                             |                                       |
| ---                                    | ---                                   |
| Declare and set new chart | `Dim oChart as Chart : Set oChart = Charts.Add`    |
|  |  |
|  |  |
|  |  |
|  |  |
|  |  |
|  |  |
|  |  |


# Snippets

## `With` Loop

Direct access to an object by using `.`:

```vb
With [OBJECT]
      [Statement]
End With
```

## Copy-Paste Range

```vb
    rng.Copy
    rng_destination.PasteSpecial xlPasteValues

    Application.CutCopyMode = False ' release the copied cells
```

## Select Case

```vb
Select [ Case ] testexpression  
    [ Case expressionlist  
        [ statements ] ]  
    [ Case Else  
        [ elsestatements ] ]  
End Select  
```

```vb
Dim number As Integer = 8
Select Case number
    Case 1 To 5
        Debug.WriteLine("Between 1 and 5, inclusive")
        ' The following is the only Case clause that evaluates to True.
    Case 6, 7, 8
        Debug.WriteLine("Between 6 and 8, inclusive")
    Case 9 To 10
        Debug.WriteLine("Equal to 9 or 10")
    Case Else
        Debug.WriteLine("Not between 1 and 10, inclusive")
End Select
```

## Text File

```vb
    Dim fileSystemObject As Object
    Dim textStream As Object
    ...
    Set fileSystemObject = CreateObject("Scripting.FileSystemObject")
    Set textStream = fileSystemObject.CreateTextFile([filename], True, False)
    textStream.Write [string]
    ...
    textStream.Close
    ' Clean up
    Set fileSystemObject = Nothing
    Set textStream = Nothing
```

## Dialog for Folder Selection

```vb
    Dim folderPath As String
    ...
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder to Save CSV Files"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With
```

## Conditional Formatting

### Cell Color 

```vb
Dim oRange as Range, oFormat as FormatCondition

Set oRange = Range(...)
set oFormat = oRange.FormatConditions.Add(xlCellValue, xlLess, "=0")
oFormat.Interior.Color = 13551615
```

```vb
With rng
    .FormatConditions.Add xlExpression, , "=AND($B3=0, $C3=0)"
    .FormatConditions(1).Interior.Color = vbYellow

### Databar

```vb
Dim oRange as Range, Dim oBar as Databar

Set oRange = Range(...)

Set oBar = oRange.FormatConditions.AddDatabar
oBar.MinPoint.Modify
    newtype:=xlConditionValueAutomaticMin
oBar.MaxPoint.Modify
    newtype:=xlConditionValueAutomaticMax
oBar.BarFillType = xlDataBarFillGradient
oBar.Direction = xlContext
oBar.NegativeBarFormat.ColorType = xlDataBarColor
oBar.BarBorder.Type = xlDataBarVoderSolid
oBar.egativeBarFormat.BorderColorType = xlDataBarColor
oBar.AxisPosition = xlDataBarAxisAutomatic
oBar.BarColor.Color = 13012579
oBar.NegativeBarFormat.Color.Color = 590255
```

## Charts

### Copy-Paste Chart

Copy a chart from a sheet and paste it as a picture into another sheet:

```vb
Set oChart = Charts.Add
oChart.SetSourceData oRange
oChart.ChartType = xlXYScatterLinesNoMarkers
Sheets(1).Select
ActiveChart.ChartArea.Copy
Sheets(2).Seelct
ActiveSheet.PasteSpecial Format:="Picture (JPEG)"
Selection.ShapeRange.ScaleWidth 0.8, msoFalse
Selection.ShapeRange.ScaleHeight 0.8, msoFalse
Selection.ShapeRange.IncrementLeft 100
Selection.ShapeRange.IncremetnTop 100
```

### Set Chart

```vb
Dim oWs as Worksheet : Set oWs = ActiveWorksheet

oWs.Shapes.AddChart2(240, x1XYScatterLines).Select
ActiveChart.SetSourceData oWs.Range(...)
ActiveChart.HasTitle = False
oWs.ChartObjects(1).Top = Range(...).Top
oWs.ChartObjects(1).Left = Range(...).Left
oWs.ChartObjects(1).Width = 300
oWs.ChartObjects(1).Height = 150
```

or into a `With` statement:

```vb
With ActiveChart
    .[property] = ...
```