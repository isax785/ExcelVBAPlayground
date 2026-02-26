# VBA TBX to Be Stored

- [VBA TBX to Be Stored](#vba-tbx-to-be-stored)
- [Undefined Actions](#undefined-actions)
- [Notes](#notes)
- [Tables](#tables)
- [Snippets](#snippets)
  - [`With` Loop](#with-loop)
  - [Cell Color Conditional Formatting](#cell-color-conditional-formatting)
  - [Databar Conditional Formatting](#databar-conditional-formatting)
  - [Set Chart](#set-chart)
  - [Select Case](#select-case)

---

> From the book `100 Excel VBA Simulations`

# Undefined Actions

- `Application.Calculation = xlCalculationManual`

# Notes

- rows and columns counter starts from `1`

# Tables

| **General**                            |                                       |
| ---                                    | ---                                   |
| Multiple statements in a single row    | *`[statement] : [statement]`*         |
| Label placed anywhere within the macro | *`[label]:`*                          |
| Jump to a specific label `[label]`     | *`GoTo [label]`*                      |
| Round to the lower integer             | `i = Int([value])`                    |
| Random integer number between 0 and N  | `Int(Rnd * (N+1))`                    |


| **Declarations**                       |                                       |
| ---                                    | ---                                   |
| Variant: can store anything            | *`Dim [name] as Variant`*             |
| Inline declarations, single type       | *`Dim [varname] as [vartype]`*        |
| Inline declaration and assignment      | *`Dim [varname] as [vartype] : [varname] = [value]`*   |
| Inline declarations, multiple types    | *`Dim [varname] as [vartype], [varname] as [vartype]`* |
| Declare and fill array    | *`Dim [arrname] as Variant : [arrname] = Array([val], [val], ...)`* |


| **Declarations**                       |                                       |
| ---                                    | ---                                   |
| Declare array of integers, undefined size | `arr() as Integer`                 |
| Re-declare array to set size           | `ReDim arr(4)`                        |
| Array as variant, then fill it         | `Dim arr as Variant : arr = Array(1, 2, 3)` |
| Size of array -> `int`                 | `UBound(arr)`                         |


| **Keywords**                           |                                        |
| ---                                    | ---                                    |
| Random value in the `[0,1]` range, `0` and `1` included     | `Rnd`             |
| Timer counting **seconds** since the midnight of the machine | `Timer`          |
| Newline keyword for string             | `vbcr`                                 |
| Prevention of system blocking (inside a loop) | `DoEvents`                      |
| Break loop                             | *`Exit [loop]`*                        |
| Row code separator                     | `:`                                    |
| Make global variable declaration (at the top of the module) | `Option Explicit` |
| Serial number of current date and time (`2/26/2026 3:23:30 PM`) | `Now`         |
|  ???                                   | `Calculate` |


| **String Manipulation**                |                                       |
| ---                                    | ---                                   |
| String concatenation (spaces are not automatically inserted) | *`"[string]" & "[string]" * [int]`* |
| Convert value to string                | *`CStr([val])`*                       |
| String to upper/lower case             | *`UCase([string])`* / *`LCase([string])`* |


| **MessageBox**                         |                                       |
| ---                                    | ---                                   |
| Open messagebox                        | *`MsgBox("[message]", [button-set])`* |
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
| Columns range                          | `Columns("A:C")`                      |
| Sorting                                | *`Range("[range-to-sort]").Sort Range("[first-sort-field]") [order]`*   |
| Sorting (`xlAscending` by default)  | `Range("A4:B270").Sort Range("B4")`   |
| Sorting descending order (??)          | `Range("D4:E270").Sort Range("E4"), xlDescending`   |


| **Cells and Ranges Formatting**        |                                          |
| ---                                    | ---                                      |
| Formatting variable                    | *`Dim [name] as FormatCondition`*        |
| Databar conditional                    | *`Dim [name] as Databar`*                |
| Cell coloring                          | *`.Cell.Interior.ColorIndex = [value]`*  |
| Range cells coloring                   | *`.Cells.Interior.ColorIndex = [value]`* |
| Thick border                           | `.borderAround, xlThick`                 |
| Horizontal alignment                   | `.Cells.HorizontalAlignment = xlCenter`  |
| Border styling                        | `.Cells.Borders.LineStyle = xlContinuous` |
| Autofit column width                   | `Cells.EntireColumn.AutoFit`             |
| Delete conditional formatting          | `Range(...).FormatConditions.Delete`     |


| **Loops and Conditionals**             |                                        | 
| ---                                    | ---                                    |
| `For` loop (`Dim i as Integer`)        | `For i = 0 to 5 ... Next i`            |
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


| **Worksheet Handling**                |                                       |
| ---                                    | ---                                   |
| Add worksheet after the active one | `Dim wks as Worksheet : Set wks = Worksheets.Add( , ActiveSheet)` |
| Assign active sheet | `Dim wks as Worksheet : Set wks = ActiveSheet` |
| Copy the active sheet | `wks.Copy, Sheets(Sheets.Count)`             |



| **Worksheet Functions**                |                                       |
| ---                                    | ---                                   |
| Call function                          | *`WorksheetFunction.[functioname(arguments)]`* |
| Count blank cells -> `int`             | *`.CountBlank([range])`*              | 
| Integer random between two values      | *`.RandBetween([int], [int])`*        |
| Count cells matchin a defined condition | *`.CountIf([range], "[condition]")`* |
| Set zoom on active window              | `ActiveWindow.Zoom = 130`             |
| Goal seek    | *`Range(...).GoalSeek [goal-value], [cell-to-change]`*  |
|  |  |
|  |  |
|  |  |
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

Direct access to a range or any other 

```vb

```

## Cell Color Conditional Formatting

```vb
Dim oRange as Range, oFormat as FormatCondition

Set oRange = Range(...)
set oFormat = oRange.FormatConditions.Add(xlCellValue, xlLess, "=0")
oFormat.Interior.Color = 13551615
```

## Databar Conditional Formatting

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

## Set Chart

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

## Select Case

