# Example: Check Row Content

- [Example: Check Row Content](#example-check-row-content)
- [1. Check if a row contains **all strings** (default behavior)](#1-check-if-a-row-contains-all-strings-default-behavior)
- [2. Check if a row contains **any** of the strings](#2-check-if-a-row-contains-any-of-the-strings)
- [3. Exact, case‑sensitive match](#3-exact-casesensitive-match)

---

Checks whether a row contains a list of strings.

```vb
Option Explicit

'
' Parameters:
'   targetRow     : Range representing a single row (any length)
'   searchTerms   : Variant array of strings to look for
'   requireAll    : If True, all terms must be found; if False, any one match is enough
'   partialMatch  : If True, substring match; if False, exact match
'   caseSensitive : If True, match is case-sensitive
'
' Returns:
'   Boolean
'
Public Function RowContainsStrings( _
    ByVal targetRow As Range, _
    ByVal searchTerms As Variant, _
    Optional ByVal requireAll As Boolean = True, _
    Optional ByVal partialMatch As Boolean = True, _
    Optional ByVal caseSensitive As Boolean = False _
) As Boolean

    Dim cell As Range
    Dim term As Variant
    Dim foundCount As Long
    Dim comparisonMethod As VbCompareMethod

    comparisonMethod = IIf(caseSensitive, vbBinaryCompare, vbTextCompare)

    For Each term In searchTerms
        Dim termFound As Boolean
        termFound = False

        For Each cell In targetRow.Cells
            If Not IsError(cell.Value) Then
                If partialMatch Then
                    If InStr(1, CStr(cell.Value), CStr(term), comparisonMethod) > 0 Then
                        termFound = True
                        Exit For
                    End If
                Else
                    If StrComp(CStr(cell.Value), CStr(term), comparisonMethod) = 0 Then
                        termFound = True
                        Exit For
                    End If
                End If
            End If
        Next cell

        If termFound Then
            foundCount = foundCount + 1
        ElseIf requireAll Then
            RowContainsStrings = False
            Exit Function
        End If
    Next term

    RowContainsStrings = (foundCount > 0)
End Function
```


# 1. Check if a row contains **all strings** (default behavior)

```vba
Dim terms As Variant
terms = Array("Pressure", "Temperature", "Flow")

If RowContainsStrings(Rows(5), terms) Then
    MsgBox "All terms found"
End If
```

***

# 2. Check if a row contains **any** of the strings

```vba
If RowContainsStrings(Rows(5), terms, requireAll:=False) Then
    MsgBox "At least one term found"
End If
```

***

# 3. Exact, case‑sensitive match

```vba
If RowContainsStrings(Rows(5), terms, True, False, True) Then
    MsgBox "Exact, case-sensitive match"
End If
```

---

[EX MOC](./ex-00_MOC.md)