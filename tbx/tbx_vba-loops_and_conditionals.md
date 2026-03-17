# VBA Toolbox - Loops and Conditionals

- [VBA Toolbox - Loops and Conditionals](#vba-toolbox---loops-and-conditionals)
  - [For](#for)
  - [With](#with)
  - [While](#while)
  - [If-Else](#if-else)
  - [Select Case](#select-case)

---


| **Loops and Conditionals**             |                                        | 
| ---                                    | ---                                    |
| `For` loop (`Dim i as Integer`)        | `For i = 0 to 5 ... Next i`            |
| `For` loop with defined step      | *`For [index] = [val] To [val] Step [val]`* |
|                                     | `For number As Double = 0 To 2 Step 0.25` |
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

## For

Syntax:

```vb
For counter [ As datatype ] = start To end [ Step step ]
    [ statements ]
    [ Continue For ]
    [ statements ]
    [ Exit For ]
    [ statements ]
Next [ counter ]
```

Standard: `For i = 1 To 6 ... Next i`

On range of selected cells: `For Each cell In rng.Cells ... Next cell`

Selected array: `For i = LBound(myArray) To UBound(myArray)`



## With

Perform a series of statements **on a specified object** without requalifying the name of the object. Direct access to an object by using `.`.

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

---

[MOC](./tbx-00_MOC.md)