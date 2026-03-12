# VBA Toolbox - Basics

- [VBA Toolbox - Basics](#vba-toolbox---basics)
- [Snippets](#snippets)
  - [Define Function](#define-function)

---

> rows and columns counter starts from `1`

| **General**                            |                                        |
| ---                                    | ---                                    |
| Multiple statements in a single row    | *`[statement] : [statement]`*          |
| Label placed anywhere within the macro | *`[label]:`*                           |
| Jump to a specific label `[label]`     | *`GoTo [label]`*                       |
| Round to the lower integer             | `i = Int([value])`                     |
| Random integer number between 0 and N  | `Int(Rnd * (N+1))`                     |
| Module -> `int`                        | *`[int] Mod [int]`*                    |
| Quotient -> `int`        | *`.Formula="=QUOTIENT([numerator], [denominator])"`* |
| Separator to write code on multiple lines | `_`                                 |
| Difference operator                    | `<>`                                   |
| Comments                        | `' comments are preceded by an apostrophe`   |


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
| Colors | `vbYellow` `vbWhite` |


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
| Cursor to wait/default        | `Application.Cursor = xlWait` or `= xlDefault` |
| Status bar text                        | *`Application.StatusBar = [str]`*     |


| **Declarations**                       | Code                                  |
| ---                                    | ---                                   |
| ilnline variable declaration and assignment | *`Dim [varName] As Integer: [varName] = value`* |
| multiple types inline decalaration     | `Dim a As Single, b As Integer`       |
| multiple single type inline declaration   | `Dim a, b, c As Double`            |

# Snippets

## Define Function

```vb
Function [funcname]([arguments])
    ...
    [funcname] = [value to be returned]
End Function
```

---

[MOC](./tbx%20-%2000%20MOC.md)