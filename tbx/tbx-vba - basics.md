# VBA Basics Toolbox

- [VBA Basics Toolbox](#vba-basics-toolbox)
- [Declarations](#declarations)
- [Ranges and Cells](#ranges-and-cells)
  - [RC Notation](#rc-notation)
  - [Names](#names)
  - [Formula](#formula)
  - [Selection](#selection)
  - [Formatting](#formatting)
- [String Manipulation](#string-manipulation)
- [Handling Sheets](#handling-sheets)
- [Loops and Conditionals](#loops-and-conditionals)
  - [For](#for)
  - [With](#with)
  - [While](#while)
  - [If-Else](#if-else)
- [Error Handling](#error-handling)

---

# Declarations

| Action                                 | Code                                  |
| ---                                    | ---                                   |
| Comments              | `' comments are preceded by an apostrophe`   |
| **Variables**  |   |
| ilnline variable declaration and assignment | *`Dim [varName] As Integer: [varName] = value`* |
| multiple types inline decalaration  | `Dim a As Single, b As Integer`  |
| multiple single type inline declaration   | `Dim a, b, c As Double` |
| **Arrays**                        |                                 |
| Array with 10 Strings             | `Dim arr(1 to 10) As String`    |
| Array with 5 Integers             | `Dim arr(0 to 4) As Integer`    |
| Array with 5 items of anything    | `Dim arr(4) As Variant`         |
| Can hold Reset to hold 10 Strings | `Dim arr() As String ReDim arr` |


# Ranges and Cells

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

## RC Notation

To be used instead of the standard cell notation with explicit row-column naming:

| The `ActiveCell` is `B11` | Result |
|---|---|
| `R1C1` | `A1` |
| `RC` | `B11` |
| `R[1]C` | `B12` |
| `R[-1]C[-1]` | `A10` |
| `=SUM(R2C:R[-1]C)` | `SUM(B2:B10)` |

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


# String Manipulation


# Handling Sheets

| Action                                 | Code                                  |
| ---                                    | ---                                   |
| Retrieve sheetname          | `sheetName = Application.Caller.Worksheet.Name`  |
| `Dim sheetName as String`   | `sheetName = ActiveSheet.Name`                   |
| Activate sheet to act with the macro | `Sheets([sheetName]).Activate`          |


# Loops and Conditionals

## For

Standard: `For i = 1 To 6 ... Next i`

On range of selected cells: `For Each cell In rng.Cells ... Next cell`

Selected array: `For i = LBound(myArray) To UBound(myArray)`

## With

Perform a series of statements **on a specified object** without requalifying the name of the object. 

```vb
With [object]
  .[property] = [val]
  .[action]
End With
```

## While

Two `while` loop statements that differ on when the condition is checked:

- `Do While [condition] ... Loop`: checked at the beginning, the loop starts only if the condition is respected.
- `Do ... Loop While [condition]`: checked at the end, the loop always starts (at least one iteration is completed).


## If-Else

Standard: `If [condition] Then ... ElseIf [condition] Then ... Else ... End If`

Single-line syntax: `If [condition] Then ... Else ... ]`


# Error Handling

| Action                                                              | Code                      |
| ------------------------------------------------------------        | -----------------------   |
| Switcher off error handling (until next _On Error_ statement)       | `On Error ()`             |
| Execution continues with the line following the error line          | `On Error Resume Next`    |
| Execution jumps to line starting with the specified label (+ colon) | `On Error GoTo [myLabel]` |
| Execution resumes with the statement that caused the error          | `Resume`                  |
| Execution resumes with the line following the error line            | `Resume Next`             |
| Execution resumes at the line starting with a specified label       | `Resume [myLabel]`        |

Example of a general error handler:

```vb
Sub AnySub()
	On Error GoTo ErrTrap
	....
	Exit Sub
	ErrTrap:
        MsgBox "Error Message"
End Sub
```