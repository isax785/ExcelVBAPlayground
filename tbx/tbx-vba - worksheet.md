# VBA Toolbox - Worksheets

- [VBA Toolbox - Worksheets](#vba-toolbox---worksheets)
- [Snippets](#snippets)
  - [Search Worksheet or Create It](#search-worksheet-or-create-it)

---

| **Handling**                           |                                       |
| ---                                    | ---                                   |
| Retrieve sheetname -> `str` | `sheetName = Application.Caller.Worksheet.Name`  |
| Activate sheet to act with the macro | `Sheets([sheetName]).Activate`          |
| Add worksheet after the active one | `Dim wks as Worksheet : Set wks = Worksheets.Add( , ActiveSheet)` |
| Assign active sheet           | `Dim wks as Worksheet : Set wks = ActiveSheet` |
| Copy the active sheet                  | `wks.Copy, Sheets(Sheets.Count)`      |
| Select sheet on number                 | `Sheet1.Select`                       |
|                                        | `Sheets(1).Select`                    |
| Worksheet name -> `str`                | *`[ws].Select`*                       |
|                                        | `If ws.Name = Sheet1.Name Then ...`   |
|                           | `If ActiveSheet.Name <> Sheet1.Name Then Exit Sub` |
| Activate worksheet                     | *`[worksheet].Activate`*              |
| **Functions**                          |                                       |
| reference to the current workbook | `ThisWorkbook`                             | 
| iterate over all the worksheets | `For Each ws In ThisWorkbook.Worksheets ... Next ws` |
| **Functions**                          |                                       |
| Call function                          | *`WorksheetFunction.[functioname(arguments)]`* |
| Count blank cells -> `int`             | *`.CountBlank([range])`*              | 
| Integer random between two values      | *`.RandBetween([int], [int])`*        |
| Count cells matchin a defined condition | *`.CountIf([range], "[condition]")`* |
| Set zoom on active window              | `ActiveWindow.Zoom = 130`             |
| Calculate average value                | *`.Average([range])`*                 |
|                                        | `.Average(Cells(5,1), Cells(1004,3))` |
| Calculate percentile vlue              | `.Percentile([range], [val])`         |
| Vertical lookup  | *`.VLookup([lookup-value], [table-array], [col-index-num], [range-lookup])`* |
| Sum whole array                        | *`.Sum([array])`*                     |
| Active window graphic properties       | `ActiveWindow.Height`                 |
|                                        | `ActiveWindow.Width`                  |
| Repeat string `n` times                | *`.Rept([str], [n])`*                 |
| Max value in an range                  | *`.Max([range])`*                     |
| Normal distribution for the specified mean and standard deviation | *`.Norm_Dist([val], [mean], [std-dev], [cumulative])`* |
| Inverse of the normal cumulative distribution for the specified mean and standard deviation | *`.Norm_Inv([probability], [mean], [std-dev])`* |
| the k-th percentile of values in a range, where k is in the range 0..1, exclusive | *`.Percentile_Exc([array], [k])`* |
| Calculate median                        | *`.Median([range])`*                 |

# Snippets

## Search Worksheet or Create It

```vb
Dim ws as Worksheet

For Each ws in Worksheets
    If ws.Name = "TargetWS" Then exists = True
Next ws

If exists = False Then
    Set ws = Worksheeets.Add(, ActiveSheet) : ws.Name = "TargetWS"
Else
    Set ws = Worksheets("TargetWS")
End If

ws.Activate
```

---

[MOC](./tbx%20-%2000%20MOC.md)