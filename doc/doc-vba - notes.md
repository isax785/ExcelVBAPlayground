# VBA - Notes

- [VBA - Notes](#vba---notes)
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
- [Sub and Functions](#sub-and-functions)
  - [Function](#function)
  - [Sub](#sub)
  - [Calling Sub and Function procedures](#calling-sub-and-function-procedures)
  - [Execute Macro on Cell Change](#execute-macro-on-cell-change)
- [Charts](#charts)
  - [Chart Axis](#chart-axis)
  - [Chart Range](#chart-range)
    - [Dynamic Range Cahnge without VBA](#dynamic-range-cahnge-without-vba)
- [Debug](#debug)
  - [Immediate Window](#immediate-window)
  - [MessageBox](#messagebox)
  - [Print](#print)
- [Simulation](#simulation)
  - [Screen Updating](#screen-updating)
  - [Cursor](#cursor)
  - [Progress Bar](#progress-bar)
  - [Sub](#sub-1)
  - [Interrupt a Sub](#interrupt-a-sub)
- [Error Handling](#error-handling)
- [Best Practice](#best-practice)
  - [Selection position](#selection-position)
- [Cookbook](#cookbook)
  - [Concatenate Range of Cells](#concatenate-range-of-cells)
  - [Hexadecimal to VBA Color](#hexadecimal-to-vba-color)
  - [Goal Seek](#goal-seek)
  - [Copy-Paste Results on Table](#copy-paste-results-on-table)
  - [Select Range and Cycle on It](#select-range-and-cycle-on-it)
  - [Delete Rows](#delete-rows)
  - [Progress Bar](#progress-bar-1)
  - [Select Cells and Highlight Cells on Another Column](#select-cells-and-highlight-cells-on-another-column)
- [References](#references)

---

Some useful snippet code to speed-up the macro implementation.

# Variables

| Data type                                | Storage size              |          | Range                                                                                                                                                                                                                                                                               |
| ---------------------------------------- | ------------------------- | -------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Byte 1 byte                              |                           | 0 to 255 |                                                                                                                                                                                                                                                                                     |
| Boolean                                  | 2 bytes                   | b        | True or False                                                                                                                                                                                                                                                                       |
| Integer                                  | 2 bytes                   | i        | -32,768 to 32,767                                                                                                                                                                                                                                                                   |
| Long (long integer)                      | 4 bytes                   | l        | -2,147,483,648 to 2,147,483,647                                                                                                                                                                                                                                                     |
| Single (single precision floating-point) | 4 bytes                   | f        | -3.402823E38 to -1.401298E-45 for negative values; 1.401298E-45 to 3.402823E38 for positive values                                                                                                                                                                                  |
| Double (double precision floating-point) | 8 bytes                   | p        | -1.79769313486231E308 to -4.94065645841247E-324 for negative values; 4.94065645841247E-324 to 1.79769313486232E308 for positive values                                                                                                                                              |
| Currency (scaled integer)                | 8 bytes                   | c        | -922,337,203,685,477.5808 to 922,337,203,685,477.5807 Decimal 14 bytes +/-79,228,162,514,264,337,593,543,950,335 with no decimal point; +/-7.9228162514264337593543950335 with 28 places to the right of the decimal; smallest non-zero number is +/-0.0000000000000000000000000001 |
| Date                                     | 8 bytes                   | d        | January 1, 100 to December 31, 9999                                                                                                                                                                                                                                                 |
| String (variable length)                 | 10 bytes + string length  | s        | 0 to approximately 2 billion                                                                                                                                                                                                                                                        |
| String (fixed-length)                    | Length of string          | s        | 1 to approximately 65,400                                                                                                                                                                                                                                                           |
| Variant  (with numbers)                  | 16 bytes                  | v        | Any numeric value up to the range of a Double                                                                                                                                                                                                                                       |
| Variant (with characters)                | 22 bytes + string  length | v        | Same range as for variable-length String                                                                                                                                                                                                                                            |

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

# Sub and Functions

## [Function](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/function-statement)

**<u>Syntax</u>**

> [**Public** | **Private** | **Friend**] [ **Static** ] **Function** *name* [ ( *arglist* ) ] [ **As** *type* ]
> [ *statements* ]
> [ *name* **=** *expression* ]
> [ **Exit Function** ]
> [ *statements* ]
> [ *name* **=** *expression* ]
> **End Function**

The **Function** statement syntax has these parts:

| Part         | Description                                                  |
| :----------- | :----------------------------------------------------------- |
| **Public**   | Optional. Indicates that the **Function** procedure is accessible to all other procedures in all [modules](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#module). If used in a module that contains an **Option Private**, the procedure is not available outside the [project](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#project). |
| **Private**  | Optional. Indicates that the **Function** procedure is accessible only to other procedures in the module where it is declared. |
| **Friend**   | Optional. Used only in a [class module](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#class-module). Indicates that the **Function** procedure is visible throughout the project, but not visible to a controller of an instance of an object. |
| **Static**   | Optional. Indicates that the **Function** procedure's local [variables](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#variable) are preserved between calls. The **Static** attribute doesn't affect variables that are declared outside the **Function**, even if they are used in the procedure. |
| *name*       | Required. Name of the **Function**; follows standard variable naming conventions. |
| *arglist*    | Optional. List of variables representing arguments that are passed to the **Function** procedure when it is called. Multiple variables are separated by commas. |
| *type*       | Optional. [Data type](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#data-type) of the value returned by the **Function** procedure; may be [Byte](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#byte-data-type), [Boolean](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#boolean-data-type), [Integer](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#integer-data-type), [Long](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#long-data-type), [Currency](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#currency-data-type), [Single](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#single-data-type), [Double](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#double-data-type), [Decimal](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#decimal-data-type) (not currently supported), [Date](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#date-data-type), [String](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#string-data-type) (except fixed length), [Object](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#object), [Variant](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#variant-data-type), or any [user-defined type](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#user-defined-type). |
| *statements* | Optional. Any group of statements to be executed within the **Function** procedure. |
| *expression* | Optional. Return value of the **Function**.                  |

The *arglist* argument has the following syntax and parts:

> [ **Optional** ] [ **ByVal** | **ByRef** ] [ **ParamArray** ] *varname* [ ( ) ] [ **As** *type* ] [ **=** *defaultvalue* ]

| Part           | Description                                                  |
| :------------- | :----------------------------------------------------------- |
| **Optional**   | Optional. Indicates that an argument is not required. If used, all subsequent arguments in *arglist* must also be optional and declared by using the **Optional** keyword. **Optional** can't be used for any argument if **ParamArray** is used. |
| **ByVal**      | Optional. Indicates that the argument is passed [by value](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#by-value). |
| **ByRef**      | Optional. Indicates that the argument is passed [by reference](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#by-reference). **ByRef** is the default in Visual Basic. |
| **ParamArray** | Optional. Used only as the last argument in *arglist* to indicate that the final argument is an **Optional** array of **Variant** elements. The **ParamArray** keyword allows you to provide an arbitrary number of arguments. It may not be used with **ByVal**, **ByRef**, or **Optional**. |
| *varname*      | Required. Name of the variable representing the argument; follows standard variable naming conventions. |
| *type*         | Optional. Data type of the argument passed to the procedure; may be **Byte**, **Boolean**, **Integer**, **Long**, **Currency**, **Single**, **Double**, **Decimal** (not currently supported) **Date**, **String** (variable length only), **Object**, **Variant**, or a specific [object type](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#object-type). If the parameter is not **Optional**, a user-defined type may also be specified. |
| *defaultvalue* | Optional. Any [constant](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#constant) or constant expression. Valid for **Optional** parameters only. If the type is an **Object**, an explicit default value can only be **Nothing**. |

<u>**Remarks**</u>

If not explicitly specified by using **Public**, **Private**, or **[Friend](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/friend-keyword)**, **Function** procedures are public by default.

If **[Static](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/static-statement)** isn't used, the value of local variables is not preserved between calls.

The **Friend** keyword can only be used in class modules. However, **Friend** procedures can be accessed by procedures in any module of a project. A **Friend** procedure does not appear in the [type library](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#type-library) of its parent class, nor can a **Friend** procedure be late bound.

**Function** procedures can be recursive; that is, they can call themselves to perform a given task. However, recursion can lead to stack overflow. The **Static** keyword usually isn't used with recursive **Function** procedures.

All executable code must be in procedures. You can't define a **Function** procedure inside another **Function**, **[Sub](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/sub-statement)**, or **Property** procedure.

The **[Exit Function](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/exit-statement)** statement causes an immediate exit from a **Function** procedure. Program execution continues with the statement following the statement that called the **Function** procedure. Any number of **Exit Function** statements can appear anywhere in a **Function** procedure.

Like a **Sub** procedure, a **Function** procedure is a separate procedure that can take arguments, perform a series of statements, and change the values of its arguments. However, unlike a **Sub** procedure, you can use a **Function** procedure on the right side of an [expression](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#expression) in the same way you use any intrinsic function, such as **Sqr**, **Cos**, or **Chr**, when you want to use the value returned by the function.

You call a **Function** procedure by using the function name, followed by the argument list in parentheses, in an expression. See the **[Call](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/call-statement)** statement for specific information about how to call **Function** procedures.

To return a value from a function, assign the value to the function name. Any number of such assignments can appear anywhere within the procedure. If no value is assigned to *name*, the procedure returns a default value: a numeric function returns 0, a string function returns a zero-length string (""), and a **Variant** function returns [Empty](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#empty). A function that returns an object reference returns **Nothing** if no object reference is assigned to *name* (using **Set**) within the **Function**.

The following example shows how to assign a return value to a function. In this case, **False** is assigned to the name to indicate that some value was not found.

VBCopy

```vb
Function BinarySearch(. . .) As Boolean 
'. . . 
 ' Value not found. Return a value of False. 
 If lower > upper Then 
  BinarySearch = False 
  Exit Function 
 End If 
'. . . 
End Function
```

Variables used in **Function** procedures fall into two categories: those that are explicitly declared within the procedure and those that are not.

Variables that are explicitly declared in a procedure (using **Dim** or the equivalent) are always local to the procedure. Variables that are used but not explicitly declared in a procedure are also local unless they are explicitly declared at some higher level outside the procedure.

A procedure can use a variable that is not explicitly declared in the procedure, but a naming conflict can occur if anything you defined at the [module level](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#module-level) has the same name. If your procedure refers to an undeclared variable that has the same name as another procedure, constant, or variable, it is assumed that your procedure refers to that module-level name. Explicitly declare variables to avoid this kind of conflict. You can use an **[Option Explicit](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/option-explicit-statement)** statement to force explicit declaration of variables.

Visual Basic may rearrange arithmetic expressions to increase internal efficiency. Avoid using a **Function** procedure in an arithmetic expression when the function changes the value of variables in the same expression. For more information about arithmetic operators, see [Operators](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/operator-summary).

<u>**Example**</u>

The name of the function corresponds to the function output and must be declared in it's first statement. Also input must be declared (see example below):

```vb
Function colSum(cell As String) As Double
    ' Computes the sum of a column of values
    Range(Range(cell), Range(cell).End(xlDown)).Select
    colSum = Application.WorksheetFunction.Sum(Selection)
End Function
```

Then it can be used as follows:

> Dim column_sum as Double
>
> column_sum = colSum("A1")

All the functions can also be used in the worksheet by writing it in a cell.

Below an example where also _Case_ is used:

```vb
Function DayName(InputDate as Date)
    ' InputDate must be a date
	Dim DayNumber As Integer
	DayNumber = Weekday(InputDate, vbSunday)
	Select Case DayNumber
		Case 1
    		DayName = "Sunday"
		Case 2
    		DayName = "Monday"
		Case 3
    		DayName = "Tuesday"
		Case 4
    		DayName = "Wednesday"    
		Case 5
    		DayName = "Thursday"
		Case 6
    		DayName = "Friday"
		Case 7
    		DayName = "Saurday"
	End Select
End Function
```

## [Sub](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/sub-statement)

**<u>Syntax</u>**

> [ **Private** | **Public** | **Friend** ] [ **Static** ] **Sub** *name* [ ( *arglist* ) ]
> [ *statements* ]
> [ **Exit Sub** ]
> [ *statements* ]
> **End Sub**

| Part         | Description                                                  |
| :----------- | :----------------------------------------------------------- |
| **Public**   | Optional. Indicates that the **Sub** procedure is accessible to all other procedures in all [modules](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#module). If used in a module that contains an **Option Private** statement, the procedure is not available outside the [project](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#project). |
| **Private**  | Optional. Indicates that the **Sub** procedure is accessible only to other procedures in the module where it is declared. |
| **Friend**   | Optional. Used only in a [class module](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#class-module). Indicates that the **Sub** procedure is visible throughout the [project](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#project), but not visible to a controller of an instance of an object. |
| **Static**   | Optional. Indicates that the **Sub** procedure's local [variables](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#variable) are preserved between calls. The **Static** attribute doesn't affect variables that are declared outside the **Sub**, even if they are used in the procedure. |
| *name*       | Required. Name of the **Sub**; follows standard [variable](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#variable) naming conventions. |
| *arglist*    | Optional. List of variables representing arguments that are passed to the **Sub** procedure when it is called. Multiple variables are separated by commas. |
| *statements* | Optional. Any group of [statements](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#statement) to be executed within the **Sub** procedure. |

The *arglist* argument has the following syntax and parts:

[ **Optional** ] [ **ByVal** | **ByRef** ] [ **ParamArray** ] *varname* [ ( ) ] [ **As** *type* ] [ **=** *defaultvalue* ]

| Part           | Description                                                  |
| :------------- | :----------------------------------------------------------- |
| **Optional**   | Optional. [Keyword](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#keyword) indicating that an argument is not required. If used, all subsequent arguments in *arglist* must also be optional and declared by using the **Optional** keyword. **Optional** can't be used for any argument if **ParamArray** is used. |
| **ByVal**      | Optional. Indicates that the argument is passed [by value](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#by-value). |
| **ByRef**      | Optional. Indicates that the argument is passed [by reference](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#by-reference). **ByRef** is the default in Visual Basic. |
| **ParamArray** | Optional. Used only as the last argument in *arglist* to indicate that the final argument is an **Optional** [array](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#array) of **Variant** elements. The **ParamArray** keyword allows you to provide an arbitrary number of arguments. **ParamArray** can't be used with **ByVal**, **ByRef**, or **Optional**. |
| *varname*      | Required. Name of the variable representing the argument; follows standard variable naming conventions. |
| *type*         | Optional. [Data type](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#data-type) of the argument passed to the procedure; may be [Byte](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#byte-data-type), [Boolean](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#boolean-data-type), [Integer](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#integer-data-type), [Long](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#long-data-type), [Currency](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#currency-data-type), [Single](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#single-data-type), [Double](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#double-data-type), [Decimal](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#decimal-data-type) (not currently supported), [Date](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#date-data-type), [String](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#string-data-type) (variable-length only), [Object](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#object), [Variant](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#variant-data-type), or a specific [object type](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#object-type). If the parameter is not **Optional**, a [user-defined type](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#user-defined-type) may also be specified. |
| *defaultvalue* | Optional. Any [constant](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#constant) or constant [expression](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#expression). Valid for **Optional** parameters only. If the type is an **Object**, an explicit default value can only be **Nothing**. |

**<u>Remarks</u>**

If not explicitly specified by using **Public**, **Private**, or **[Friend](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/friend-keyword)**, **Sub** procedures are public by default.

If **[Static](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/static-statement)** isn't used, the value of local variables is not preserved between calls.

The **Friend** keyword can only be used in class modules. However, **Friend** procedures can be accessed by procedures in any module of a project. A **Friend** procedure doesn't appear in the [type library](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#type-library) of its parent class, nor can a **Friend** procedure be late bound.

**Sub** procedures can be recursive; that is, they can call themselves to perform a given task. However, recursion can lead to stack overflow. The **Static** keyword usually is not used with recursive **Sub** procedures.

All executable code must be in [procedures](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#procedure). You can't define a **Sub** procedure inside another **Sub**, **Function**, or **Property** procedure.

The **Exit Sub** keywords cause an immediate exit from a **Sub** procedure. Program execution continues with the statement following the statement that called the **Sub** procedure. Any number of **Exit Sub** statements can appear anywhere in a **Sub** procedure.

Like a **Function** procedure, a **Sub** procedure is a separate procedure that can take arguments, perform a series of statements, and change the value of its arguments. However, unlike a **Function** procedure, which returns a value, a **Sub** procedure can't be used in an expression.

You call a **Sub** procedure by using the procedure name followed by the argument list. See the **[Call](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/call-statement)** statement for specific information about how to call **Sub** procedures.

Variables used in **Sub** procedures fall into two categories: those that are explicitly declared within the procedure and those that are not. Variables that are explicitly declared in a procedure (using **Dim** or the equivalent) are always local to the procedure. Variables that are used but not explicitly declared in a procedure are also local unless they are explicitly declared at some higher level outside the procedure.

A procedure can use a variable that is not explicitly declared in the procedure, but a naming conflict can occur if anything you defined at the [module level](https://docs.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#module-level) has the same name. If your procedure refers to an undeclared variable that has the same name as another procedure, constant or variable, it is assumed that your procedure is referring to that module-level name. To avoid this kind of conflict, explicitly declare variables. You can use an **Option Explicit** statement to force explicit declaration of variables.

**<u>Note</u>** : You can't use **GoSub**, **GoTo**, or **Return** to enter or exit a **Sub** procedure.

## Calling Sub and Function procedures

[Source](https://docs.microsoft.com/en-us/office/vba/language/concepts/getting-started/calling-sub-and-function-procedures)

## Execute Macro on Cell Change

**Event Handlers** are not stored in your typical module location. They are actually stored inside either your Workbook or Worksheet object. To get to the "coding area" of either your workbook or worksheet, you simply double-click **ThisWorkbook** or the sheet name (respectively) within your desired VBA Project hierarchy tree (within the **Project Window** of your Visual Basic Editor).

On Worksheet set up the *Event Handler* as follows:

- *Objects* to *Worksheet*
- *Procedure* to *Change*

![VBA Event Handler Trigger Macro Code](img_md/VBA+Event+Handler+Trigger+Macro+Code)

<img src="img_md/image-20201105115953394.png" alt="image-20201105115953394" style="zoom: 80%;" />

Then paste the following code:

```vb
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim KeyCells As Range

' The variable KeyCells contains the cells that will
    ' cause an alert when they are changed.
    Set KeyCells = Range("A1:C10")   'desired range to be monitored'

If Not Application.Intersect(KeyCells, Range(Target.Address)) _
           Is Nothing Then

' Display a message when one of the designated cells has been 
        ' changed.
        ' Place your code here.
        MsgBox "Cell " & Target.Address & " has changed."

End If
End Sub
```

# Charts

Some useful resources:

- https://peltiertech.com/Excel/Charts/index.html
  - https://peltiertech.com/Excel/Charts/axes.html

## [Chart Axis](https://exceloffthegrid.com/chart-axis-min-mix/)

Set chart axis min and max based on a cell value.

- Updates automatically whenever data changes
- Does not require user interaction – i.e. no button clicking, but updates automatically when the worksheet recalculates
- Easily portable between different worksheets

```vb
Function setChartAxis(sheetName As String, chartName As String, MinOrMax As String, _
    ValueOrCategory As String, PrimaryOrSecondary As String, Value As Variant)

'Create variables
Dim cht As Chart
Dim valueAsText As String

'Set the chart to be controlled by the function
Set cht = Application.Caller.Parent.Parent.Sheets(sheetName) _
    .ChartObjects(chartName).Chart

'Set Value of Primary axis
If (ValueOrCategory = "Value" Or ValueOrCategory = "Y") _
    And PrimaryOrSecondary = "Primary" Then

    With cht.Axes(xlValue, xlPrimary)
        If IsNumeric(Value) = True Then
            If MinOrMax = "Max" Then .MaximumScale = Value
            If MinOrMax = "Min" Then .MinimumScale = Value
        Else
            If MinOrMax = "Max" Then .MaximumScaleIsAuto = True
            If MinOrMax = "Min" Then .MinimumScaleIsAuto = True
        End If
    End With
End If

'Set Category of Primary axis
If (ValueOrCategory = "Category" Or ValueOrCategory = "X") _
    And PrimaryOrSecondary = "Primary" Then

    With cht.Axes(xlCategory, xlPrimary)
        If IsNumeric(Value) = True Then
            If MinOrMax = "Max" Then .MaximumScale = Value
            If MinOrMax = "Min" Then .MinimumScale = Value
        Else
            If MinOrMax = "Max" Then .MaximumScaleIsAuto = True
            If MinOrMax = "Min" Then .MinimumScaleIsAuto = True
        End If
    End With
End If

'Set value of secondary axis
If (ValueOrCategory = "Value" Or ValueOrCategory = "Y") _
    And PrimaryOrSecondary = "Secondary" Then

    With cht.Axes(xlValue, xlSecondary)
        If IsNumeric(Value) = True Then
            If MinOrMax = "Max" Then .MaximumScale = Value
            If MinOrMax = "Min" Then .MinimumScale = Value
        Else
            If MinOrMax = "Max" Then .MaximumScaleIsAuto = True
            If MinOrMax = "Min" Then .MinimumScaleIsAuto = True
        End If
    End With
End If

'Set category of secondary axis
If (ValueOrCategory = "Category" Or ValueOrCategory = "X") _
    And PrimaryOrSecondary = "Secondary" Then
    With cht.Axes(xlCategory, xlSecondary)
        If IsNumeric(Value) = True Then
            If MinOrMax = "Max" Then .MaximumScale = Value
            If MinOrMax = "Min" Then .MinimumScale = Value
        Else
            If MinOrMax = "Max" Then .MaximumScaleIsAuto = True
            If MinOrMax = "Min" Then .MinimumScaleIsAuto = True
        End If
    End With
End If

'If is text always display "Auto"
If IsNumeric(Value) Then valueAsText = Value Else valueAsText = "Auto"

'Output a text string to indicate the value
setChartAxis = ValueOrCategory & " " & PrimaryOrSecondary & " " _
    & MinOrMax & ": " & valueAsText

End Function
```

Then on the spreadsheet:	![image-20201027165254402](img_md/image-20201027165254402.png)

**<u>Example</u>**

![image-20201127094538738](img_md/image-20201127094538738.png)

## Chart Range

Change chart formula:

```vb
Sub ChangeSeriesFormula()
    ''' Just do active chart
    If ActiveChart Is Nothing Then
        '' There is no active chart
        MsgBox "Please select a chart and try again.", vbExclamation, _
            "No Chart Selected"
        Exit Sub
    End If

    Dim OldString As String, NewString As String, strTemp As String
    Dim mySrs As Series

    OldString = InputBox("Enter the string to be replaced:", "Enter old string")

    If Len(OldString) > 1 Then
        NewString= InputBox("Enter the string to replace " & """" _
            & OldString & """:", "Enter new string")
        '' Loop through all series
        For Each mySrs In ActiveChart.SeriesCollection
            strTemp = WorksheetFunction.Substitute(mySrs.Formula, _
                OldString, NewString)
            mySrs.Formula = strTemp
        Next
    Else
        MsgBox "Nothing to be replaced.", vbInformation, "Nothing Entered"
    End If
End Sub
```

Assuming that you want to expand the range (by adding one extra column) to add one more observation for each series in you diagram (and not to add a new series), you could use this code:

```vb
Sub ChangeChartRange()
    Dim i As Integer, r As Integer, n As Integer, p1 As Integer, p2 As Integer, p3 As Integer
    Dim rng As Range
    Dim ax As Range

    'Cycles through each series
    For n = 1 To ActiveChart.SeriesCollection.Count Step 1
        r = 0

        'Finds the current range of the series and the axis
        For i = 1 To Len(ActiveChart.SeriesCollection(n).Formula) Step 1
            If Mid(ActiveChart.SeriesCollection(n).Formula, i, 1) = "," Then
                r = r + 1
                If r = 1 Then p1 = i + 1
                If r = 2 Then p2 = i
                If r = 3 Then p3 = i
            End If
        Next i


        'Defines new range
        Set rng = Range(Mid(ActiveChart.SeriesCollection(n).Formula, p2 + 1, p3 - p2 - 1))
        Set rng = Range(rng, rng.Offset(0, 1))

        'Sets new range for each series
        ActiveChart.SeriesCollection(n).Values = rng

        'Updates axis
        Set ax = Range(Mid(ActiveChart.SeriesCollection(n).Formula, p1, p2 - p1))
        Set ax = Range(ax, ax.Offset(0, 1))
        ActiveChart.SeriesCollection(n).XValues = ax

    Next n
End Sub
```

### Dynamic Range Cahnge without VBA

You don't need VBA to make the chart dynamic. Just create dynamic named ranges that grow and shrink with the data. Your VBA for the chart can refer to that named range, without adding burden to the code. But you may not even need VBA at all. The chart defined with dynamic ranges will update instantaneously. No code required.

Range name and formula for chart labels:

```
chtLabels =Sheet1!A2:Index(Sheet1!$A:$A,counta(Sheet1!$A:$A))
```

Range name and formula for column B

```vb
chtB =offset(chtlabels,0,1)
```

Range and formula for column D

```vb
chtD =offest(chtlabels,0,3)
```

Edit the data source and instead of the fixed cell ranges, enter the named ranges in the format

```
=Sheet1!*RangeName*
```

You need to supply the respective range name to the series values and the chart's category axis.

Note: when you supply dynamic range names as the source for a chart, you MUST include either the file name or the sheet name with the range name reference. When you close and re-open the dialog, you will find that Excel automatically converts your entry to the format `[Filename]RangeName`

Note 2: there are many different formula options to create dynamic range names. In this case, we're using and index of column A and determine the last populated cell by counting the cells. This only works if all cells in column A have text. If your data has gaps in column A (which I don't think you do), different formulas can be applied to determine the range.

# Debug



## Immediate Window

Select **View -> Immediate Window** to activate. Any valid expression can be evaluated in the Immediate Window.

In the window it is possible to use the _Print method_ as follows:

> print [items] [;]

otherwise, the _print_ command can be replaced by a question mark (_?_) as follows:

> ? [item]

## MessageBox

VBA _MsgBox_ is a function that generates a dialog window.

The syntax is the following:

> MsgBox(prompt[, buttons] [, title] [, helpfile, context])

where:

- _prompt_ is the message to be visualized in the box.

- _buttons_ are the **constants** of the argument that allow to choose the type of message box:

  | Constant | Value | Description |
  | -------- | :---: | ----------- |
  |vbOKOnly|0|_ok_ button|
  |VbOKCancel|1|_ok_ and _cancel_ buttons|
  |VbAbortRetryIgnore|2|_cancel_, _retry_ and _ignore_ buttons|
  |VbYesNoCancel|3|_yes_, _no_ and _cancel_ buttons|
  |VbYesNo|4| _yes_ and _no_ buttons |
  |VbRetryCancel|5| _retry_ and _cancel_ buttons |
  |VbCritical|16| critical message |
  |VbQuestion|32| reboot request |
  |VbExclamation|48| advising message |
  |VbInformation|64| information message |

  When a **constant** is used, the instruction can be assigned to a variable (that must be declared as String or  Variant (Dim) depending on the output of the function). The instruction will be

  > Dim variable
  >
  > variable = MsgBox([message], [constant])
  
- The length of a message is 1024 characters. Special commands can be used

  | Command | Alternative Command | What It Does                          |
  | ------- | ------------------- | ------------------------------------- |
  | vbCr    | Chr(13)             | go to next line                       |
  | vbLf    | Chr(10)             | empty line                            |
  | vbCrLf  | Chr(13) + Chr(10)   | go to next line and add an empty line |

  Below an example on how to write a multi-line message:

  ```vb
  Sub Messaggio()
  	Dim Lem As String
  	Lem = Lem & "Questo è un esempio :" & vbLf & vbLf
  	Lem = Lem & "L'Autore è:" & Chr(13) & vbLf
  	Lem = Lem & "Pinco Pallino" & vbLf
  	Lem = Lem & "Questo Programma" & vbLf
  	Lem = Lem & "è tutelato dai" & vbLf & vbLf
  	Lem = Lem & "diritti d'Autore !!" & vbLf
  	Lem = Lem & "(non esiste, è falso)" & vbLf & vbLf
  	Lem = Lem & "del Codice VBa di questo messaggio" & vbLf
  	Lem = Lem & "si raccomanda di fare tutte" & vbLf
  	Lem = Lem & "le variazioni che volete." & vbLf
  	Lem = Lem & "l'Autore lo concede (?!?!?)" & vbLf
  	Lem = Lem & "Ha......Ha......Ha"
  	MsgBox Lem
  End Sub
  ```

Some examples of MsgBox messages:

```vb
MsgBox "You're in the wrong sheet!" & vbNewLine & "Please switch to the sheet: " & sheetName, vbCritical

MsgBox "OK!" & vbCrLf, vbInformation

```

Operations when closing a file (with _Cases_):

```vb
Sub Prova4()
	Dim iRisposta As Integer
	iRisposta = MsgBox("STAI PER USCIRE, VUOI SALVARE IL FILE ???", vbYesNoCancel)
	Select Case iRisposta 'impostiamo il Select Case con riferimento al messaggio restituito dalla variabile iRisposta
	Case vbYes 'se risponderemo "Si" :
		ThisWorkbook.Save 'salveremo il file e
		Application.Quit 'chiuderemo cartella ed Excel
	Case vbCancel 'se sceglieremo "Annulla":
		Exit Sub 'usciremo dalla routine
	Case vbNo 'se sceglieremo "No":
		Application.Quit 'chiuderemo cartella ed Excel senza salvare
	Case Else
		End Select
End Sub
```

Set a predefined value (_vbNo_) to not accomplish an action:

> Cancel = (MsgBox("Sicuro di voler chiudere la finestra ?", vbYesNo) = vbNo)

## Print

The _Print method_ sends output to the immediate window whenever the _Debug object_ prefix is included:

> Debug.Print [items] [;]

This string can be written in any row of the code.

# Simulation

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

# Error Handling

| Command                 | Action                                                       |
| ----------------------- | ------------------------------------------------------------ |
| On Error ()             | Switcher off error handling (until next _On Error_ statement) |
| On Error Resume Next    | Execution continues with the line following the error line   |
| On Error GoTo *myLabel* | Execution jumps to line starting with the specified label (+ colon) |
| Resume                  | Execution resumes with the statement that caused the error   |
| Resume Next             | Execution resumes with the line following the error line     |
| Resume _myLabel_        | Execution resumes at the line starting with a specified label |

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



# Best Practice

Some guidelines to improve the simulation experience.

## Selection position

At the end of a macro, it could be useful to place the cell selection in a specific cell in order to avoid to move manually when the simulation ends. For example:

> Range("A1").Select   'selection of the worksheet top-left cell

The same can be done to move to a specific sheet:

> Sheets(sheetName).Activate
>
> Range("A1").Select

# Cookbook

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



# References

- 300 examples: https://www.excel-easy.com/examples.html
