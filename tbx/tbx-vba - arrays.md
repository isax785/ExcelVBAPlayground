# VBA ToolBox - Arrays


| **Arrays**                        |                                 |
| ---                                    | ---                                   |
| Array with 10 Strings             | `Dim arr(1 to 10) As String`    |
| Array with 5 Integers             | `Dim arr(0 to 4) As Integer`    |
| Array with 5 items of anything    | `Dim arr(4) As Variant`         |
| Can hold Reset to hold 10 Strings | `Dim arr() As String ReDim arr` |
| Variant: can store anything            | *`Dim [name] as Variant`*             |
| Inline declarations, single type       | *`Dim [varname] as [vartype]`*        |
| Inline declaration and assignment      | *`Dim [varname] as [vartype] : [varname] = [value]`*   |
| Inline declarations, multiple types    | *`Dim [varname] as [vartype], [varname] as [vartype]`* |
| Declare and fill array    | *`Dim [arrname] as Variant : [arrname] = Array([val], [val], ...)`* |
| Declare array of integers, undefined size | `arr() as Integer`                 |
| Declare array of `N-M+1` integers      | `Dim arr(M to N) as Integer`          |
| Declare array of `N-M+1` anything      | `Dim arr(M to N) as Variant`          |
| Resize array to set size (5)           | `ReDim arr(4)`                        |
| Resize array lower-upper bound         | *`ReDim [arr]([lower] to [upper])`*   |
|                                        | `ReDim arr(0 to 10)`                  |
| Resize array size (upper bound only) without changing the contained data | `ReDim Preserve arr(10)` |
| Array as variant, then fill it         | `Dim arr as Variant : arr = Array(1, 2, 3)` |
| Array of range addresses | `Dim arr as Variant : arr = Array("E1:F13", "H3:I13")` |
| Size of 1D array -> `int`              | `UBound(arr)`                         |
| Size of multidimensional array         | *`UBound([arr], [dim])`*              |


---

[MOC](./tbx%20-%2000%20MOC.md)