# Sub and Functions

- [Sub and Functions](#sub-and-functions)
  - [Functions](#functions)
  - [Subs](#subs)
  - [Calling Sub and Function procedures](#calling-sub-and-function-procedures)
  - [Execute Macro on Cell Change](#execute-macro-on-cell-change)

---

## Functions

[Docs.Microsoft - Function](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/function-statement)

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

## Subs

[Docs.Microsoft - Sub](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/sub-statement)

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
