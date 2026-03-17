# VBA Toolbox - Cells and Ranges Toolbox

- [VBA Toolbox - Cells and Ranges Toolbox](#vba-toolbox---cells-and-ranges-toolbox)
  - [Names](#names)
  - [Formula](#formula)
  - [Selection](#selection)
  - [Formatting](#formatting)
- [Snippets](#snippets)
  - [Ranges](#ranges)
    - [Copy-Paste Range](#copy-paste-range)
    - [Set Range From Itself](#set-range-from-itself)
    - [Set Range from Another Range](#set-range-from-another-range)
  - [Conditional Formatting](#conditional-formatting)
    - [Cell Color](#cell-color)
    - [Bars](#bars)

---

> Reference: `Cells(nrRow, nrCol)`, `nrRow` and `nrCol` start from `1`.

| Action           | Code                                  |
| ---              | ---                                   |
| Declare range    | `Dim rng as Range`                    |
| cell `A1`        | `rng = Range("A1")`                   |
| cell `A1:A7`     | `rng = Range("A1:A7")`                |
| cell `A1 + A5`   | `rng = Range("A1,A5")`                |
| cell `A1:A5`     | `rng = Range("A1","A5")`              |
| cell `A1`        | `rng = Cells(1, 1)`                   |
| cell `B5`        | `rng = Cells(5, 2)`                   |
| cell `All Cells` | `rng = Cells()`                       |
| cell `A1`        | `rng = Range("A1:A5").Cells(1, 1)`    |
| cell `B1`        | `rng = Range("B1:C5").Cells(1, 1)`    |
| cell `A5`        | `rng = Range("A1:A5").Cells(5, 1)`    |
| cell `A6`        | `rng = Range("A5").Range("A2") [OR: Offset(1,0)]` |
| cell `A1:A5`     | `rng = Range(Cells(1, 1), Cells(5, 1))` |
| cell `A1:C5`     | `rng = Range(Cells(1, 1), Cells(5, 3)`  |
| cell `B2`        | `rng = Range(Cells(2, 2), Cells(5, 5)).Range("A1")` |
| cell "named"     | `rng = Range("named")`                |
| check empty cell | `IsEmpty(Range(...).Value)`           |

**RC Notation** : To be used instead of the standard cell notation with explicit row-column naming:

| The `ActiveCell` is `B11` | Result |
|---|---|
| `R1C1` | `A1` |
| `RC` | `B11` |
| `R[1]C` | `B12` |
| `R[-1]C[-1]` | `A10` |
| `=SUM(R2C:R[-1]C)` | `SUM(B2:B10)` |


| **Declarations**                   |                                       |
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
|                   | `Range(...).Formula = "=[formula]" & [strintg/value] & "[formula]` | 
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
| Assign values to a range with formula  | `Range("D12:N12").Formula = Range("D12:N12").Value` |
| Resize range (both parameters are optional) | *`rng.Resize([RowSize], ColumnSize)`* |
| Delete region                      | *`[range].CurrentRegion.Delete [option]`* |
|                                        | `[option] = xlShiftToLeft, xlShiftUp` |
| Range adaptive formula, i.e. it changes with the cell on the whole column | `range.Columns(4).Formula = "Sum(A2:C2)"` |
| Sorting [doc](https://learn.microsoft.com/en-us/office/vba/api/excel.range.sort)| *`[range].Sort [Key1], [Order1], [Key2], [Type], [Order2], [Key3], [Order3], [Header], [OrderCustom], [MatchCase], [Orientation], [SortMethod], [DataOption1], [DataOption2], DataOpt[ion3`* |
|   | `Columns("A:C").Sort key1:=Range("C2"), order1:=xlAscending, header:=xlYes` |
|                                   | `.Sort oSort, xlAscending, , , , , , xlYes` |
| Assign name (i.e. named range)         | *`[range].Name = [name]`*              |
|                                        | `rng.Name = "data"`                    |
| Delete name from a sheet               | *`[sheet].Names([name]).Delete`*       |
|                                        | `Sheet1.Names("data").Delete`          |
| Hide/Show row                          | `.EntireRow.Hidden = True` / `= False` |
| Hide/Show column                    | `.EntireColumn.Hidden = True` / `= False` |
| Formula array                          | *`[range].FormulaArray = "[formula]"`* |
|                          | `Range("C2:C10").FormulaArray = "=A2:A10=""hello"""` |

| **Formatting**        |                                          |
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
| Percent formatting                     | `rng = FormatPercent([val], [decimals])` |
| Number formatting                      | *`[rng].NumberFormat = [format]`*        |
|                                    | `Cells.EntireColumn.NumberFormat = "0.0000"` |
|                                    | `Cells.EntireColumn.NumberFormat = "m/d/yy"` |
|                                | `Cells.EntireColumn.NumberFormat = ""[h]:mm:ss"` |
| Number format decimals                    | *`FormatNumber([value], [decimals])`* |
| Add conditional formatting to range   | `rng.FormatConditions(xlExpression, xlFormula, "=B2<" & value)` |
| Coloring with RGB (max 255)           | `rng.Interior.Color = RGB(0, 0, 0)`       |
| Format currency with decimals         | *`FormatCurrency([value], [decimals])`*   |
| Format number with decimals           | *`FormatNumber([value], [decimals])`*     |
| Draw cell thick border                | *`[range].BorderAround , xlThick`*        |
| Set bold font                         | *`[range].Cells.Font.Bold = True`*        |
| Percentage and conditional color  | `.Cells.NumberFormat = "0.00%;[Red] -0.00%"`  |
| Borders weight (bottom) | `Range(Cells(2, 1), Cells(1, 8)).Borders(xlEdgeBottom).Weight = xlMedium` |
| Borders weight (right) | `Range(Cells(2, 1), Cells(8, 1)).Borders(xlEdgeRight).Weight = xlMedium` |

## Names

Use cell names inside the macro: after the name has been assigned to the cell, refer to it with `Range([name])`.

## Formula

Write formula with:
- **standard notation**: `Worksheets("Sheet1").Range("A5").Formula = "=A4+A10"`
- **RC notation**: `Worksheets("Sheet1").Range("A5").FormulaR1C1 = "=R4C1+R10C1"`

Note: `"=A4+A10"` and `"=R4C1+R10C1"` are the same cell references.

## Selection

Operations to be carried out on **selection**:

| Action                                 | Code                                  |
| ---                                    | ---                                   |
| Range selection in the **active** sheet | `Range([cellName]).Select`            | 
| Range selection from **another** sheet| `Sheets([sheetName]).Range([cellName]).Select` |
| Sum all the cells             | `Application.WorksheetFunction.Sum(Selection)` |
| Clear the cell contents                | `Selection.ClearContents`             |
| Clear the cell format                  | `Selection.ClearFormats`              |
| Clear the cell comments                | `Selection.ClearComments`             |
| Clear the cell color                   | `Selection.Interior.color = xlNone`   |
| Clear all the cell                     | `Selection.Clear`                     |
| Select multiple cells        | *`Range(Range([cellName]), Range([cellName]).End([direction])).Select`*`*` |
|   | *`Range(Range([cellName]), Range([cellName]).Offset([nrRows], [nrCols]))`*    |
| Select a region                        | `Range([cellName]).CurrentRegion.Select` |
| Select multiple rows                   | `Rows("2:38").Select`                    | 
| Select multiple columns                | `Columns("B:D").Select`                  |
| Hide/show selection row/column | `Selection.EntireRow.Hidden = True 'hide` |
|                                | `Selection.EntireRow.Hidden = False 'show` |
|                                | `Selection.EntireColumn.Hidden = True 'hide` |
|                                | `Selection.EntireColumn.Hidden = False 'show` |

`*`: `[direction]` = `xlUp`, `xlDown`, `xltoLeft`, or `xltoRight`.

## Formatting

The cell format can be set as follows **before writing the value** in the cell:

- `colSum = Format(Application.WorksheetFunction.Sum(Selection), "0.00")`
- `colSum = Format(Application.WorksheetFunction.Sum(Selection), "0.00%")`

# Snippets

## Ranges

### Copy-Paste Range

```vb
    rng.Copy
    rng_destination.PasteSpecial xlPasteValues  ' values only

    Application.CutCopyMode = False ' release the copied cells
```


### Set Range From Itself

Set a range from itself:

```vb
Dim oRange as Range
Set oRange = Range("B6").CurrentRegion

With oRange
    Set oRange = .Offset(1, 0).Resize(.Rows.Count-1, .Columns.Count-1)
End With
```

### Set Range from Another Range

Set a range from another range by resizing:

```vb
Dim oRange, oTable as Range

Set oRange = Range("B14").CurrentRegion
With oRange
    Set oTable = .Offset(1, 1).Resize(.Rows.Count-1, .Columns.Count-1)
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

### Bars

```vb
With oRange.Columns(5)
    Dim oBar as Databar
    .Select
    Set oBar = Selection.FormatConditions.AddDatabar
    oBar.MinPoint.Modify newtype:=xlConditionValueAutomaticMin
    oBar.MaxPoint.Modify newtype:=xlConditionValueAutomaticMax
    oBar.BarFillType = xlDataBarFillGradient
    oBar.Direction = xlContext
    oBar.NegativeBarFormat.ColorType = xlDataBarColor
    oBar.BarBorder.Type = xlDataBarBorderSolid
    oBar.NegativeBarFormat.BorderColorType = xlDataBarColor
    oBar.AxisPosition = xlDAtaBarAxisAutomatic
End With
```

---

[MOC](./tbx-00_MOC.md)