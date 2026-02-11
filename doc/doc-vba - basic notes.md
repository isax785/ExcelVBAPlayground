# VBA - Basic Notes

- [VBA - Basic Notes](#vba---basic-notes)
- [Variables](#variables)
  - [Arrays](#arrays)
    - [Jagged Array](#jagged-array)
  - [Comments](#comments)
- [Sheet](#sheet)
- [Ranges and Cells](#ranges-and-cells)
  - [Ranges vs. Cells](#ranges-vs-cells)
  - [Formula R1C1](#formula-r1c1)
  - [Copy and Paste](#copy-and-paste)
  - [Delete Row of a Selected Cell](#delete-row-of-a-selected-cell)
- [Loops and Conditional](#loops-and-conditional)
  - [_For_ Loop](#for-loop)
  - [_While_ Loop](#while-loop)
  - [_Do_ Loop](#do-loop)
  - [*If Else*](#if-else)
  - [*With*](#with)
  - [Useful Checks](#useful-checks)
- [Names](#names)

---

Some useful snippet code to speed-up the macro implementation.

# Variables

| Data type       | Storage size    |   | Range   |
| --- | --- | --- | --- |
| Byte 1 byte | | 0 to 255 | |
| Boolean | 2 bytes | b | True or False |
| Integer      | 2 bytes                   | i        | -32,768 to 32,767   |
| Long (long integer)                      | 4 bytes                   | l        | -2,147,483,648 to 2,147,483,647                 |
| Single (single precision floating-point) | 4 bytes                   | f        | -3.402823E38 to -1.401298E-45 for negative values; 1.401298E-45 to 3.402823E38 for positive values       |
| Double (double precision floating-point) | 8 bytes                   | p        | -1.79769313486231E308 to -4.94065645841247E-324 for negative values; 4.94065645841247E-324 to 1.79769313486232E308 for positive values|
| Currency (scaled integer)                | 8 bytes                   | c        | -922,337,203,685,477.5808 to 922,337,203,685,477.5807 Decimal 14 bytes +/-79,228,162,514,264,337,593,543,950,335 with no decimal point; +/-7.9228162514264337593543950335 with 28 places to the right of the decimal; smallest non-zero number is +/-0.0000000000000000000000000001 |
| Date         | 8 bytes                   | d        | January 1, 100 to December 31, 9999             |
| String (variable length)                 | 10 bytes + string length  | s        | 0 to approximately 2 billion                    |
| String (fixed-length)                    | Length of string          | s        | 1 to approximately 65,400                       |
| Variant  (with numbers)                  | 16 bytes                  | v        | Any numeric value up to the range of a Double   |
| Variant (with characters)                | 22 bytes + string  length | v        | Same range as for variable-length String        |

Declare a variable and assign it in the same row:

`Dim [varName] As Integer: [varName] = 0`

```vb
Dim i as Integer: i = 100
```

Declare multiple variables on the same row

```vb
Dim a As Single, b As Single, c As Single, x As Double, y As Double, i As Integer
```

> In VBA, the following declaration
>
> ```vb
> Dim a, b, c As Single
> ```
>
> It is equivalent to this:
>
> ```vb
> Dim a As Variant, b As Variant, c As Single
> ```

## Arrays


| Declaration                   | Result                                |
| :---------------------------- | ------------------------------------- |
| Dim arr(1 to 10) As String    | Can hold **10 Strings**               |
| Dim arr(0 to 4) As Integer    | Can hold **5 Integers**               |
| Dim arr(4) As Variant         | Can hold **5 items of anything**      |
| Dim arr() As String ReDim arr | Can hold **Reset to hold 10 Strings** |

### Jagged Array

| (0)  | (1)    | (2)    | (3)    | (4)    | (5)    |
| ---- | ------ | ------ | ------ | ------ | ------ |
| (0)  | (1)(0) | (2)(0) | (3)(0) | (4)(0) | (5)(0) |
| (0)  | (1)(1) | (2)(1) |        | (4)(1) | (5)(1) |
| (0)  |        | (2)(2) |        | (4)(2) | (5)(2) |
| (0)  |        | (2)(3) |        |        | (5)(3) |


## Comments

The comment text must be preceded by an apostrophe:

> 'this is the comment text

To comment/uncomment multiple lines together use the **Comment Block** button (the bar with the button in the VBA window should be activated).

# Sheet

Retrieve the name of the current sheet:

>`sheetNameVar = Application.Caller.Worksheet.Name`

or

>` sheetNameVar = ActiveSheet.Name`

Activate a sheet to allow the macro working on it:

>`Sheets([sheetName]).Activate`

Below an example:

```vb
Dim thisSheet as String

thisSheet = Application.Caller.Worksheet.Name

Sheets("anotherSheet").Activate
'then come back to the previous sheet
Sheets(thisSheet).Activate
```

# Ranges and Cells

Selection of :  https://docs.microsoft.com/en-us/office/troubleshoot/office-developer/select-cells-rangs-with-visual-basic 

Select a cell with _Range_:  `Range([cellName]).Select`

or with _Cell_:  `Cells(nrRow, nrCol)`

> _nrRow_ and _nrCol_ start from **1**.

The selection from another sheet can be done without activating the sheet: `Sheets([sheetName]).Range([cellName]).Select`

Write a formula in a cell or cell range:  `Range("A1:A12").Formula = "=Rand()"`

It is like writing the formula as string inside the cell, standard Excel function syntax must be followed.

Check whether a cell is empty (returns _TRUE_) or not (returns _FALSE_): `IsEmpty(Range([cellName]).Value)`

The cell format can be set as follows before writing the value in a cell: 
- `colSum = Format(Application.WorksheetFunction.Sum(Selection), "0.00")`
- `colSum = Format(Application.WorksheetFunction.Sum(Selection), "0.00%")`

Select a range of cells by starting from _[cellName]_ and following a direction till the end of the cells with a value:  `Range(Range([cellName]), Range([cellName]).End([direction])).Select`

Where _[direction] = xlUp, xlDown, xltoLeft, xltoRight_ depending on the direction that the selection _ must follow.

Another method is by considering an offset: `Range(Range([cellName]), Range([cellName]).Offset([nrRows], [nrCols]))`

Select a region: `Range([cellName]).CurrentRegion.Select`

Select multiple rows/columns: `Rows("2:38").Select`  and `Columns("B:D").Select`

Hide/show selection row/column:

> `Selection.EntireRow.Hidden = True    'hide`
>
> `Selection.EntireRow.Hidden = False   'show`
>
> `Selection.EntireColumn.Hidden = True    'hide`
>
> `Selection.EntireColumn.Hidden = False   'show`

The two methods can be mixed, for example to select a table as in the following example code:

```vb
Sub clearAmbTable()
    Dim startAmb As String
    startAmb = "B8"
    If Not IsEmpty(Range(startAmb).Value) Then
        Range(Range(startAmb), Range(startAmb).End(xlDown).offset(0, 6)).Select
        Selection.ClearContents
    End If
    Range("A1").Select
End Sub
```

The range must be declared before assign it to a variable:

> Dim rng as Range

Once the cell range is selected, it is possible to 

- **sum** all the cells:

  > Application.WorksheetFunction.Sum(Selection)

- **clear** the cell _contents_:

  > Selection.ClearContents

- **clear** the cell _format_:

  > Selection.ClearFormats

- **clear** the cell _comments_:

  > Selection.ClearComments

- **clear** the cell _color_:

  > Selection.Interior.color = xlNone

- **clear** _all_ the cell:

  > Selection.Clear

Note that **_Selection_** is a valid keyword that can be used to make actions on selected cells without assigning them to a variable.

## Ranges vs. Cells

| Command                                     | Which Cell |
| ------------------------------------------- | ---------- |
| Range("A1")                                 | A1         |
| Range("A1:A7")                              | A1:A7      |
| Range("A1,A5")                              | A1 + A5    |
| Range("A1","A5")                            | A1:A5      |
| Cells(1, 1)                                 | A1         |
| Cells(5, 2)                                 | B5         |
| Cells()                                     | All Cells  |
| Range("A1:A5").Cells(1, 1)                  | A1         |
| Range("B1:C5").Cells(1, 1)                  | B1         |
| Range("A1:A5").Cells(5, 1)                  | A5         |
| Range("A5").Range("A2")  [OR: Offset(1,0)]  | A6         |
| Range(Cells(1, 1), Cells(5, 1))             | A1:A5      |
| Range(Cells(1, 1), Cells(5, 3))             | A1:C5      |
| Range(Cells(2, 2), Cells(5, 5)).Range("A1") | B2         |

## Formula R1C1

FormulaR1C1 has the same behavior as Formula, only using R1C1 style annotation, instead of A1 annotation. In A1 annotation you would use:

```vb
Worksheets("Sheet1").Range("A5").Formula = "=A4+A10"
```

In R1C1 you would use:

```vb
Worksheets("Sheet1").Range("A5").FormulaR1C1 = "=R4C1+R10C1"
```

| The ActiveCell is B11 | Result      |
| --------------------- | ----------- |
| R1C1                  | A1          |
| RC                    | B11         |
| R[1]C                 | B12         |
| R[-1]C[-1]            | A10         |
| =SUM(R2C:R[-1]C)      | SUM(B2:B10) |

```vb
Sub Macro1()
	Range("D4").Select
	ActiveCell.FormulaR1C1 = "=R[-1]c[-2]*10"
	Range("D5").Select
End Sub
```

## Copy and Paste

Simple example to copy and paste value and/or format:

```vb
Worksheets(1).Cells(i, 3).Copy
Worksheets(2).Cells(a, 15).PasteSpecial Paste:=xlPasteFormats
Worksheets(2).Cells(a, 15).PasteSpecial Paste:=xlPasteValues
```

Three Methods to Copy & Paste with VBA, source: https://www.excelcampus.com/vba/copy-paste-cells-vba-macros/

```vb
Sub Range_Copy_Examples()
'Use the Range.Copy method for a simple copy/paste
    'The Range.Copy Method - Copy & Paste with 1 line
    Range("A1").Copy Range("C1")
    Range("A1:A3").Copy Range("D1:D3")
    Range("A1:A3").Copy Range("D1")
    'Range.Copy to other worksheets
    Worksheets("Sheet1").Range("A1").Copy Worksheets("Sheet2").Range("A1")
    'Range.Copy to other workbooks
    Workbooks("Book1.xlsx").Worksheets("Sheet1").Range("A1").Copy _
        Workbooks("Book2.xlsx").Worksheets("Sheet1").Range("A1")
End Sub

Sub Paste_Values_Examples()
'Set the cells' values equal to another to paste values
    'Set a cell's value equal to another cell's value
    Range("C1").Value = Range("A1").Value
    Range("D1:D3").Value = Range("A1:A3").Value     
    'Set values between worksheets
    Worksheets("Sheet2").Range("A1").Value = Worksheets("Sheet1").Range("A1").Value     
    'Set values between workbooks
    Workbooks("Book2.xlsx").Worksheets("Sheet1").Range("A1").Value = _
        Workbooks("Book1.xlsx").Worksheets("Sheet1").Range("A1").Value        
End Sub

Sub PasteSpecial_Examples()
'Use the Range.PasteSpecial method for other paste types
    'Copy and PasteSpecial a Range
    Range("A1").Copy
    Range("A3").PasteSpecial Paste:=xlPasteFormats    
    'Copy and PasteSpecial a between worksheets
    Worksheets("Sheet1").Range("A2").Copy 
    Worksheets("Sheet2").Range("A2").PasteSpecial Paste:=xlPasteFormulas    
    'Copy and PasteSpecial between workbooks
    Workbooks("Book1.xlsx").Worksheets("Sheet1").Range("A1").Copy
    Workbooks("Book2.xlsx").Worksheets("Sheet1").Range("A1").PasteSpecial Paste:=xlPasteFormats    
    'Disable marching ants around copied range
    Application.CutCopyMode = False
End Sub
```

## Delete Row of a Selected Cell

```vb
Sub find_delete_row()
    Dim r As Long
    'The panel Find must be open
    Cells.Find(What:="TImestamp", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        
    Debug.Print ActiveCell.Row
    r = ActiveCell.Row
    Rows(r).Select
        Selection.Delete Shift:=xlUp   'Shift up the cells below
End Sub
```



# Loops and Conditional

- **Single** loop

  ```vb
  For i = 1 To 6
  	Cells(i, 1).Value = 100
  Next i
  ```

- **Double** loop

  ```vb
  For i = 1 To 6
  	For j = 1 To 2
  		Cells(i, j).Value = 100
  	Next j
  Next i
  ```

- **Triple** loop

  ```vb
  For c = 1 To 3
  	For i = 1 To 6
  		For j = 1 To 2
  			Worksheets(c).Cells(i, j).Value = 100
  		Next j
  	Next i
  Next c
  ```

- **Do While** loop

  ```vb
  Dim i As Integer
  i = 1
  Do While i < 6
  	Cells(i, 1).Value = 20
  	i = i + 1
  Loop
  ```

  ```vb
  Do While Cells(i, 1).Value <> ""
  	Cells(i, 2).Value = Cells(i, 1).Value + 10
  	i = i + 1
  Loop
  ```



## _For_ Loop

A _for_ cycle can be run on a **selected range of cells**. The range must be declared, and at the end of the _for_ cycle the command _Next_ must be written.

```vb
Dim rng As Range

Range(Range([cellID]), Range([cellID]).End(xlDown)).Select
Set rng = Selection

For Each cell In rng.Cells
	debug.print cell.value
	Next cell
```

For a **selected array.**

```vb
For i = LBound(myArray) To UBound(myArray)
msg = msg & myArray(i) & vbNewLine 'display in a message box
 	Next i
```

## _While_ Loop

Syntax:

> Do While [condition]
>
> ​	--- actions ---
>
> Loop

## _Do_ Loop

Exit from the loop with the _End_ command, otherwise with the _Loop Until_ condition:

```vb
Do
    Range(seek_1).GoalSeek goal:=Range(goal_1), ChangingCell:=Range(bychng_1)

    If Abs(Range(bychng_1).Value) > maxBychng_1 Then
        nrIter = nrIter + 1
        If nrIter > maxIter Then
        	Exit Do
        End If
    End If
Loop Until Abs(Range(seek_1) - Range(goal_1)) < 0.01 And nrIter <= maxIter
```

## [*If Else*]( https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/if-then-else-statement)

```vb
' Multiline syntax:
If condition [ Then ]
    [ statements ]
[ ElseIf elseifcondition [ Then ]
    [ elseifstatements ] ]
[ Else
    [ elsestatements ] ]
End If

' Single-line syntax:
If condition Then [ statements ] [ Else [ elsestatements ] ] 
```

## *With*

The **With** statement allows you to perform a series of statements on a specified object without requalifying the name of the object. For example, to change a number of different properties on a single object, place the property assignment statements within the **With** control structure, referring to the object once instead of referring to it with each property assignment.

```vbscript
Sub FormatRange() 
 With Worksheets("Sheet1").Range("A1:C10") 
 .Value = 30 
 .Font.Bold = True 
 .Interior.Color = RGB(255, 255, 0) 
 End With 
End Sub
```

You can nest **With** statements by placing one **With** block within another. However, because members of outer **With** blocks are masked within the inner **With** blocks, you must provide a fully qualified object reference in an inner **With** block to any member of an object in an outer **With** block.

```vbscript
Sub MyInput() 
 With Workbooks("Book1").Worksheets("Sheet1").Cells(1, 1) 
 .Formula = "=SQRT(50)" 
 With .Font 
 .Name = "Arial" 
 .Bold = True 
 .Size = 8 
 End With 
 End With 
End Sub
```

## Useful Checks

- whether a cell is empty

  ```vbscript
  IsEmpty(Range(cellName))=True '| False depending on the needs
  ```

# Names

Use cell names inside the macro: after the name has been assigned to the cell, refer to it by using _Range_. Below an example with the names _SUM_LEN_1_ and _L_W_1_:

```vb
Function min_w(depth As Long)
	If depth < Range("SUM_LEN_1").Value Then
    	min_w = Range("L_W_1").Value * 0.868 * depth / 1000
	Else
    	min_w = Range("L_W_1").Value * 0.868 * Range("SUM_LEN_1").Value / 1000 + Range("L_W_2").Value * 0.868 * (depth - Range("SUM_LEN_1").Value) / 1000
	End If
End Function
```

Another example looping into a range of cells:

```vb
Sub ApplyColor() 
    Const Limit As Integer = 25 
    For Each c In Range("MyRange") 
        If c.Value > Limit Then 
            c.Interior.ColorIndex = 27 
        End If 
    Next c 
End Sub
```



