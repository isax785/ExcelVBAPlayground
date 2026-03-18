# VBA Regex

- [VBA Regex](#vba-regex)
- [Parse String `A18370`](#parse-string-a18370)
  - [1) Simple validator (exact match)](#1-simple-validator-exact-match)
  - [2) Finder inside longer text (extract the first or all matches)](#2-finder-inside-longer-text-extract-the-first-or-all-matches)
  - [3) Make it configurable (optional)](#3-make-it-configurable-optional)
  - [Summary](#summary)
  - [Notes](#notes)
- [Date Manipulation](#date-manipulation)
    - [What it does](#what-it-does)
  - [VBA Code](#vba-code)
  - [How to use](#how-to-use)
    - [Examples](#examples)
  - [Notes \& tips](#notes--tips)

---


# Parse String `A18370`

If by “strings like `A18370`” you mean **one uppercase letter followed by five digits**, you can detect them reliably in VBA using a regular expression.

Below is a small, dedicated UDF (user-defined function) you can use in cells or in your VBA procedures.

**Patterns**

```vb
' first letter 'A', then 1 to 10 numbers 
    .pattern = "\bA[0-9]{1,10}\b" 
' first any upper-case letter, then 5 numbers 
    .Pattern = "^[A-Z][0-9]{5}$"
```

***

## 1) Simple validator (exact match)

**Returns `TRUE` only if the entire string is exactly one letter + five digits (e.g., `A18370`).**

```vb
' Requires reference to "Microsoft VBScript Regular Expressions 5.5"
' (VBA Editor → Tools → References…)

Public Function IsCodeAxxxxx(ByVal text As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    
    With re
        .Pattern = "^[A-Z][0-9]{5}$"
        .Global = False
        .IgnoreCase = False
    End With
    
    IsCodeAxxxxx = re.Test(text)
End Function
```

**Worksheet usage:**

```excel
=IsCodeAxxxxx(A2)
```

**VBA usage:**

```vb
If IsCodeAxxxxx("A18370") Then
    Debug.Print "Valid"
End If
```

***

## 2) Finder inside longer text (extract the first or all matches)

**Returns the first occurrence of the pattern inside any text (empty string if none).**

```vb
Public Function FindCodeAxxxxx(ByVal text As String) As String
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    
    With re
        .Pattern = "\b[A-Z][0-9]{5}\b"
        .Global = False
        .IgnoreCase = False
    End With
    
    If re.Test(text) Then
        Set m = re.Execute(text)(0)
        FindCodeAxxxxx = m.Value
    Else
        FindCodeAxxxxx = ""
    End If
End Function
```

**Worksheet usage:**

```excel
=FindCodeAxxxxx(A2)
```

**Example:**

*   Text: `"Ref: PO A18370 due next week"` → Result: `A18370`

**To return **all** matches (comma-separated):**

```vb
Public Function FindAllCodeAxxxxx(ByVal text As String) As String
    Dim re As Object, matches As Object, m As Object
    Dim out As String
    Set re = CreateObject("VBScript.RegExp")
    
    With re
        .Pattern = "\b[A-Z][0-9]{5}\b"
        .Global = True
        .IgnoreCase = False
    End With
    
    If re.Test(text) Then
        Set matches = re.Execute(text)
        For Each m In matches
            If Len(out) > 0 Then out = out & ", "
            out = out & m.Value
        Next m
    End If
    
    FindAllCodeAxxxxx = out
End Function
```

***

## 3) Make it configurable (optional)

If your pattern might change (e.g., allow lowercase, allow two letters, or a different digit count), you can parameterize it:

```vb
' pattern example:
'   "^[A-Z][0-9]{5}$"    -> exact match, one uppercase + five digits
'   "\b[A-Z]{1,2}\d{5}\b"-> one or two letters + five digits, word-bounded, find inside text

Public Function RegexTest(ByVal text As String, ByVal pattern As String, _
                          Optional ByVal ignoreCase As Boolean = False, _
                          Optional ByVal globalFind As Boolean = False) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    
    With re
        .Pattern = pattern
        .Global = globalFind
        .IgnoreCase = ignoreCase
    End With
    
    RegexTest = re.Test(text)
End Function
```

***

## Summary

| Function              | Purpose                                             | Pattern Used        | Typical Use                               |
| --------------------- | --------------------------------------------------- | ------------------- | ----------------------------------------- |
| `IsCodeAxxxxx`        | Validate that the **entire** string is `A`+5 digits | `^[A-Z][0-9]{5}$`   | Cell formula data validation              |
| `FindCodeAxxxxx`      | Extract the **first** occurrence inside text        | `\b[A-Z][0-9]{5}\b` | Pull the first code from notes/remarks    |
| `FindAllCodeAxxxxx`   | Extract **all** occurrences (comma-separated)       | `\b[A-Z][0-9]{5}\b` | Consolidate multiple codes from a cell    |
| `RegexTest` (generic) | Test any custom regex pattern                       | (caller-supplied)   | Reuse for future variations of the format |

***

## Notes

*   For best performance and intellisense, set a reference to **Microsoft VBScript Regular Expressions 5.5** (VBA Editor → *Tools* → *References*).
*   If you cannot set the reference, the late-binding `CreateObject("VBScript.RegExp")` shown above already works.

If your actual pattern differs (e.g., two letters + five digits, allow lowercase, require a specific leading letter, etc.), tell me the exact rules and I’ll tailor the function.

# Date Manipulation

Below is a robust Excel VBA **worksheet function** that parses “loose” date strings like:

*   `12-Dec-24 A`
*   `14-Jan-26`
*   `14-Jan-26*`

It tolerates trailing annotations (e.g., `*`, `A`, comments), variable delimiters (`-`, `/`, space, `.`), and month names (short or long). It returns a real Excel date or `#VALUE!` if it can’t parse.

### What it does

*   **Strips trailing non‑date annotations** (e.g., `*`, `A`, comments after a space).
*   Accepts **`d-MMM-yy`**, **`dd-MMM-yyyy`**, and also **numeric** formats like `dd/mm/yy`.
*   Uses **`DateSerial`** with month-name mapping (avoids locale pitfalls from `CDate`).
*   Handles **2‑digit years** with a configurable **pivot year** (default: `30` → `00–30` maps to 2000–2030; `31–99` maps to 1931–1999).

***

## VBA Code

> Paste this into a **standard module** (e.g., `Module1`) in the VBA editor.

```vb
Option Explicit

' Public UDF: use in cells like =ParseLooseDate(A1)
' pivotYear: 2-digit years <= pivotYear map to 2000+, otherwise 1900+
Public Function ParseLooseDate(ByVal txt As String, Optional ByVal pivotYear As Integer = 30) As Variant
    Dim cleaned As String, dt As Date
    Dim rx As Object, m As Object
    Dim d As Long, y As Long, mon As Long
    
    On Error GoTo EH
    
    cleaned = Trim$(txt)
    If Len(cleaned) = 0 Then GoTo FAIL
    
    ' 1) Try: extract d-[mon]-y with month name (short or long), allowing dividers or spaces
    ' Examples matched: 12-Dec-24, 14 Jan 2026, 1.March.25, 07 Feb-2024, etc.
    Set rx = CreateObject("VBScript.RegExp")
    With rx
        .Pattern = "^\s*(\d{1,2})[\s\-/\.]+([A-Za-z]{3,9})[\s\-/\.]+(\d{2,4})"
        .IgnoreCase = True
        .Global = False
    End With
    
    If rx.Test(cleaned) Then
        Set m = rx.Execute(cleaned)(0)
        d = CLng(m.SubMatches(0))
        mon = MonthFromName(CStr(m.SubMatches(1)))
        If mon = 0 Then GoTo NEXT_PATTERN
        
        y = NormalizeYear(CStr(m.SubMatches(2)), pivotYear)
        If y = 0 Then GoTo NEXT_PATTERN
        
        If IsValidDMY(d, mon, y) Then
            ParseLooseDate = DateSerial(y, mon, d)
            Exit Function
        Else
            GoTo NEXT_PATTERN
        End If
    End If
    
NEXT_PATTERN:
    ' 2) Try: numeric-only formats dd-mm-yy[yy], dd/mm/yy, dd mm yyyy, etc.
    ' We will also tolerate trailing annotations like * or letters.
    With rx
        .Pattern = "^\s*(\d{1,2})[\s\-/\.]+(\d{1,2})[\s\-/\.]+(\d{2,4})"
    End With
    
    If rx.Test(cleaned) Then
        Set m = rx.Execute(cleaned)(0)
        d = CLng(m.SubMatches(0))
        mon = CLng(m.SubMatches(1))
        y = NormalizeYear(CStr(m.SubMatches(2)), pivotYear)
        If mon < 1 Or mon > 12 Then GoTo FAIL
        If IsValidDMY(d, mon, y) Then
            ParseLooseDate = DateSerial(y, mon, d)
            Exit Function
        Else
            GoTo FAIL
        End If
    End If
    
    ' 3) Last resort: strip trailing non-date chars and try DateValue
    cleaned = KeepUntilNonDateCore(cleaned)
    If Len(cleaned) > 0 Then
        On Error Resume Next
        dt = DateValue(cleaned)
        If Err.Number = 0 Then
            ParseLooseDate = dt
            Exit Function
        End If
        On Error GoTo EH
    End If
    
FAIL:
    ParseLooseDate = CVErr(xlErrValue)
    Exit Function
    
EH:
    ParseLooseDate = CVErr(xlErrValue)
End Function

' Map month names to month numbers; supports short and long names, case-insensitive.
Private Function MonthFromName(ByVal monTxt As String) As Integer
    Dim t As String
    t = LCase$(Trim$(monTxt))
    
    ' Accept 3+ chars; resolve ambiguous (e.g., "Ma" would be ambiguous) – we expect >=3 chars
    Select Case True
        Case Left$(t, 3) = "jan": MonthFromName = 1
        Case Left$(t, 3) = "feb": MonthFromName = 2
        Case Left$(t, 3) = "mar": MonthFromName = 3
        Case Left$(t, 3) = "apr": MonthFromName = 4
        Case Left$(t, 3) = "may": MonthFromName = 5
        Case Left$(t, 3) = "jun": MonthFromName = 6
        Case Left$(t, 3) = "jul": MonthFromName = 7
        Case Left$(t, 3) = "aug": MonthFromName = 8
        Case Left$(t, 3) = "sep": MonthFromName = 9
        Case Left$(t, 3) = "oct": MonthFromName = 10
        Case Left$(t, 3) = "nov": MonthFromName = 11
        Case Left$(t, 3) = "dec": MonthFromName = 12
        Case Else: MonthFromName = 0
    End Select
End Function

' Convert 2-digit year to 4-digit by pivot rule; pass-through for 4-digit.
' Returns 0 if invalid.
Private Function NormalizeYear(ByVal yTxt As String, ByVal pivotYear As Integer) As Long
    Dim n As Long
    yTxt = Trim$(yTxt)
    If Len(yTxt) = 0 Then Exit Function
    If Not IsNumeric(yTxt) Then Exit Function
    n = CLng(yTxt)
    
    Select Case Len(yTxt)
        Case 1, 2
            If n <= pivotYear Then
                NormalizeYear = 2000 + n
            Else
                NormalizeYear = 1900 + n
            End If
        Case 3
            ' Unusual; treat as invalid
            NormalizeYear = 0
        Case Else
            If n >= 1000 And n <= 9999 Then
                NormalizeYear = n
            Else
                NormalizeYear = 0
            End If
    End Select
End Function

' Validate that day/month/year make a real date (e.g., 31-Apr invalid)
Private Function IsValidDMY(ByVal d As Long, ByVal m As Long, ByVal y As Long) As Boolean
    On Error GoTo BAD
    If y < 100 Or y > 9999 Then GoTo BAD
    If m < 1 Or m > 12 Then GoTo BAD
    If d < 1 Or d > 31 Then GoTo BAD
    
    Dim test As Date
    test = DateSerial(y, m, d)
    ' check the components match (DateSerial will auto-rollover)
    If Year(test) = y And Month(test) = m And Day(test) = d Then
        IsValidDMY = True
    Else
        IsValidDMY = False
    End If
    Exit Function
BAD:
    IsValidDMY = False
End Function

' Keep the substring from the start up to the last character that still looks like date core.
' Allowed early: digits, letters, space, -, /, .
Private Function KeepUntilNonDateCore(ByVal s As String) As String
    Dim i As Long, ch As String, lastGood As Long
    lastGood = 0
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "0" And ch <= "9") _
        Or (ch >= "A" And ch <= "Z") _
        Or (ch >= "a" And ch <= "z") _
        Or ch = " " Or ch = "-" Or ch = "/" Or ch = "." Then
            lastGood = i
        Else
            Exit For
        End If
    Next i
    If lastGood > 0 Then
        KeepUntilNonDateCore = Trim$(Left$(s, lastGood))
    Else
        KeepUntilNonDateCore = ""
    End If
End Function
```

***

## How to use

*   In a cell, type:  
    `=ParseLooseDate(A1)`  
    or with a different 2‑digit year pivot:  
    `=ParseLooseDate(A1, 40)`  ➜ `00–40` → 2000–2040; `41–99` → 1941–1999.

### Examples

| Input           | Result (Date) |
| --------------- | ------------- |
| `12-Dec-24 A`   | 12-Dec-2024   |
| `14-Jan-26`     | 14-Jan-2026   |
| `14-Jan-26*`    | 14-Jan-2026   |
| `01/03/24 note` | 01-Mar-2024   |
| `7.Feb.2025*`   | 07-Feb-2025   |

> If a string cannot be parsed, the function returns **`#VALUE!`**.

***

## Notes & tips

*   This UDF uses **late-bound RegExp** (`CreateObject("VBScript.RegExp")`), so you don’t need to set a VBA reference.
*   If you expect US-style numeric inputs (`MM/DD/YY`), you could add a **second numeric pattern** branch or add a parameter to select the order. Right now the numeric branch assumes **`DD/MM/YY`** (day first), consistent with your examples.
*   If you anticipate month names in languages other than English, we can extend `MonthFromName` with those names.

Would you like me to add a parameter to switch between **DD/MM** and **MM/DD** numeric parsing, or auto-detect based on values?

---

[EX MOC](./ex-00_MOC.md)