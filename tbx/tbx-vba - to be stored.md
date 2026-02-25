# VBA TBX to Be Stored

- [VBA TBX to Be Stored](#vba-tbx-to-be-stored)
- [Notes](#notes)
- [Table](#table)
- [Snippets](#snippets)
  - [`With` Loop](#with-loop)

---

> From the book `100 Examples...`

# Notes

- rows and columns counters start from 1

# Table

| **General**                            |                                       |
| ---                                    | ---                                   |
| Multiple statements in a single row    | *`[statement] : [statement]`*         |
| Random value in the `[0,1]` range, `0` and `1` included     | `Rnd`            |
| Round to the lower integer             | `i = Int([value])`                    |
| Jump to a specific label `[label]`     | *`GoTo [label]`*                      |
| Label placed anywhere within the macro | *`[label]:`*                          |
| String concatenation (spaces are not automatically inserted) | *`"[string]" & "[string]" * [int]`* |
| Newline keyword for string             | `vbcr`                                |
| **Types**                              |                                       |
| Variant: can store anything            | *`Dim [name] as Variant`*             |
|  |  |
|  |  |

| **Messagebox**                         |                                       |
| ---                                    | ---                                   |
| Open messagebox                        | *`MsgBox("[message]", [button-set])`* |
| Messagebox button set                  | `vbOkCancel`, `vbYesNoCancel`|
| Buttons signals                        | `vbOK`, `vbCancel`, `vbYes`, `vbNo`   |
| Conditional | *`If MsgBox("[message]", [button-set]) = [signal] Then [action] `* |
| Get messagebox output | `Dim msg as Variant` `msg = MsgBox(...)`               |
| oputput cases                          | `Yes` -> `Case 6`                     |
|                                        | `No` -> `Case 7`                      |
|                                        | `Cancel` -> `Case 2`                  |

| **Declarations**                       |                                       |
| ---                                    | ---                                   |
| Inline declarations, single type | *`Dim [varname] as [vartype]`* |
| Inline declaration and assignment | *`Dim [varname] as [vartype] : [varname] = [value]`* |
| Inline declarations, multiple types | *`Dim [varname] as [vartype], [varname] as [vartype]`* |
| Declare and fill array | *`Dim [arrname] as Variant : [arrname] = Array([val], [val], ...)`* |
|  |  |

| **Formatting**                         |                                       |
| ---                                    | ---                                   |
| cell coloring                        | *`.Cell.Interior.ColorIndex = [value]`* |
|  |  |
|  |  |

| **Cells and Ranges**                   |                                       |
| ---                                    | ---                                   |
| Offset (`[row]` and `[col]` are incremental values) | `Range(...).Offset([row], [col])`     |
| Get address -> `str`                   | *`[range].Address`*                   |
| Select region | `Range(...).CurrentRegion`  |
| Region row and column count -> `in  t` | `.CurrentRegion.Rows.Count` `.CurrentRegion.Columns.Count`  |
| Clear the region | `.CurrentRegion.Delete` | 
| Access to cell value with coordinates | *`Cells([row], [col])`*  |

| **Loops and Conditionals**             |                                       | 
| ---                                    | ---                                   |
| `For` loop (`Dim i as Integer`)        | `For i = 0 to 5 ... Next i`           |
| `For` loop break                       | `Exit For`                            |
| `With` loop (`Dim r as Range : Set r = Range(...)`) | `With r ... .Cells(1, 1) = ... End With` |
| `If` condition                       | *`If [condition] Then ... End If`*       |
| `If ... Else` condition           | *`If [condition] Then ... Else ... End If`* |
| `If ... ElseIf ... Else` condition | *`If [condition] Then ... ElseIf [condition] Then ... Else ... End If`* |
| Inline conditional assignment          | *`i = IIF([condition], [true-return], [false-return])`* |
| `Do` loop - a break condition is mandatory  | *`Do ... Loop`*  |
| Break `Do` loop             | `Exit Do` | 
| `Do` loop with breking condition  | *`Do Until [condition] ... Loop`*           |
|                                   | *`Do ... Loop Until [condition]`*           |
| Select cases | *`Select case [var] Case [val]: ... Case [val]: ... End Select`* |

| **Worksheet Functions**                |                                       |
| ---                                    | ---                                   |
| Call function                          | *`WorksheetFunction.[functioname(arguments)]`* |
| Count blank cells -> `int`             | *`.CountBlank([range])`*              | 
|  |  |
|  |  |
|  |  |
|  |  |
|  |  |
|  |  |
|  |  |
|  |  |
|  |  |
|  |  |
|  |  |
|  |  |
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
