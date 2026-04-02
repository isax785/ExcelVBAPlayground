# VBA Toolbox - Strings


| **Manipulation**                |                                       |
| ---                                    | ---                                   |
| String length -> `int`                 | *`Len([str])`*                        |
| String concatenation (spaces are not automatically inserted) | *`"[string]" & "[string]" * [int]`* |
| Convert value to string                | *`CStr([val])`*                       |
| String to upper/lower case             | *`UCase([string])`* / *`LCase([string])`* |
| Character to integer                   | `CInt(...)`                           |
| Reverse a string -> `str`              | *`[str_rev] = StrReverse([str])`*     |
| Left characters                        | *`Left([str], [n_char])`*             |
| Right characters                       | *`Right([str], [n_char])`*            |
| Trim left, right, both-sides leading spaces | *`LTrim([string])` `RTrim([string])` `Trim([string])`* |
| Trim with high efficiency`*`           | *`Trim$([string])`* |
| Compare strings: partial match         | *`InStr(1, CStr([value]), CStr([value]), vbBinaryCompare)`* |
| Compare strings: complete match        | *`StrComp(CStr([value]), CStr([value]), vbTextCompare)`* |

`*` it works directly on strings instead of variants

## String Comparison

Reference: [Example: Check Row Content](../ex/ex_vba_check_row_content.md)

```vb
Function CompareStrings (
    ...,
    Optional ByVal partialMatch As Boolean = True, _
    Optional ByVal caseSensitive As Boolean = False _
) As Boolean
ong
    Dim comparisonMethod As VbCompareMethod

    comparisonMethod = IIf(caseSensitive, vbBinaryCompare, vbTextCompare)
    ...
        If partialMatch Then
            If InStr(1, CStr(cell.Value), CStr(term), comparisonMethod) > 0 Then
                ...
            End If
        Else
            If StrComp(CStr(cell.Value), CStr(term), comparisonMethod) = 0 Then
                ...
            End If
```

---

[MOC](./tbx-00_MOC.md)
