# VBA Regex Guide

Below is a compact, VBA‑focused “regex starter kit” you can drop into your Excel projects. It includes:

1.  A **markdown table** summarizing the main regex actions in VBA and the corresponding code patterns.
2.  A **catalog of significant, real‑world regex application examples** (with patterns and VBA snippets).

> **Note (VBA):** Excel uses the `VBScript.RegExp` engine. It supports groups, lookaheads, and lookbehinds (since Win10 builds; older systems lacked lookbehind). It **does not** support inline flags (`(?i)`), atomic groups `(?>)`, or named groups `(?<name>...)`. Use `.IgnoreCase` and `.Global` properties for flags.

***

## 1) Main regex actions in Excel VBA (with code)

> *Tip:* Set reference **Tools → References → Microsoft VBScript Regular Expressions 5.5** for early binding (IntelliSense). Otherwise, late binding with `CreateObject("VBScript.RegExp")` works too.

### Quick helper (optional)

```vb
' Late-bound helper to create a configured RegExp
Private Function Rx(pattern As String, Optional ignoreCase As Boolean = False, Optional globalFind As Boolean = False) As Object
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    With re
        .Pattern = pattern
        .IgnoreCase = ignoreCase
        .Global = globalFind
    End With
    Set Rx = re
End Function
```

### Summary table

| Action  | What it does  | Core members  | Minimal VBA code  |  |  |
| --- | --- | --- | --- | --- | --- |
| **Test (boolean)**  | Check if text matches a pattern (anywhere or exact via anchors) | `.Test(text)`  | `Dim re As Object: Set re = Rx("^[A-Z]\\d{5}$")`</br>`If re.Test("A18370") Then Debug.Print "Match"`  |  |  |
| **Execute (iterate matches)**   | Get all matches and groups  | `.Execute(text)` → `Matches` → `Match.SubMatches`  | `Dim re As Object, m As Object`</br>`Set re = Rx("(\\d{4})-(\\d{2})-(\\d{2})", False, True)`</br>`For Each m In re.Execute("On 2025-10-31 and 2026-01-01")`</br>`Debug.Print m.Value, m.SubMatches(0), m.SubMatches(1), m.SubMatches(2)`</br>`Next` |  |  |
| **Replace (transform text)**    | Substitute text based on a pattern                              | `.Replace(text, replacement)` (supports backrefs like `$1`) | `Dim re As Object: Set re = Rx("(\\d{4})-(\\d{2})-(\\d{2})")`</br>`Debug.Print re.Replace(\"2025-10-31\", \"$3/$2/$1\") '31/10/2025`  |  |  |
| **Capture groups**  | Extract specific subparts  | Parentheses `( ... )`, accessed via `Match.SubMatches(i)`   | `Dim re As Object, m As Object`</br>`Set re = Rx("([A-Z])(\\d{5})")`</br>`Set m = re.Execute(\"A18370\")(0)`</br>`Debug.Print m.SubMatches(0) 'A`</br>`Debug.Print m.SubMatches(1) '18370`  |  |  |
| **Word boundaries & anchors**   | Constrain position (start/end/word)  | `^`, `$`, `\b`, `\B`  | `Set re = Rx(\"\\b[A-Z]\\d{5}\\b\")`</br>`Debug.Print re.Test(\"Ref A18370 due\") 'True`  |  |  |
| **Lookarounds**  | Match with context without consuming it  | `(?=...)`, `(?!...)`, `(?<=...)`, `(?<!...)`  | `' Find ABC only when followed by -123`</br>`Set re = Rx(\"ABC(?=-123)\")`</br>`Debug.Print re.Test(\"ABC-123\") 'True`  |  |  |
| **Quantifiers**  | Control repetition  | `?`, `*`, `+`, `{m,n}`  | `Set re = Rx(\"[A-Z]{2}\\d{4,6}\")`</br>`Debug.Print re.Test(\"AB183700\") 'True`  |  |  |
| **Character classes & escapes** | Sets and types  | `[A-Z] [0-9] \d \w \s`, negation `[^...]`  | `Set re = Rx(\"[A-Z]\\d{5}\")` |  |  |
| **Alternation & grouping**      | Choice between subpatterns  | `A  \| B`, grouping `( ... )`  | `Set re = Rx("(CAT \| DOG)-\d+")` | | |

***

## 2) Significant regex applications (patterns + VBA snippets)

> **All examples** use late binding via `CreateObject("VBScript.RegExp")` or the `Rx` helper above. Adjust `.IgnoreCase` and `.Global` as needed.

### A. Validate codes / IDs

**One uppercase letter + five digits (e.g., `A18370`)**

```vb
Dim re As Object: Set re = Rx("^[A-Z]\d{5}$")
Debug.Print re.Test("A18370") 'True
```

**Two letters + 6–8 digits (e.g., `PO12345678`)**

```vb
Set re = Rx("^[A-Z]{2}\d{6,8}$")
```

**Hex string (8 or 16 hex chars)**

```vb
Set re = Rx("^(?:[0-9A-F]{8}|[0-9A-F]{16})$")
```

***

### B. Find & extract from free text

**First code inside text (`A` + 5 digits)**

```vb
Dim m As Object: Set re = Rx("\b[A-Z]\d{5}\b")
If re.Test(Range("A2").Value) Then
    Set m = re.Execute(Range("A2").Value)(0)
    Range("B2").Value = m.Value
End If
```

**All codes (comma‑separated)**

```vb
Function FindAllCodes$(ByVal s$)
    Dim re As Object, ms As Object, m As Object, out$
    Set re = Rx("\b[A-Z]\d{5}\b", False, True)
    If re.Test(s) Then
        Set ms = re.Execute(s)
        For Each m In ms
            If Len(out) > 0 Then out = out & ", "
            out = out & m.Value
        Next
    End If
    FindAllCodes = out
End Function
```

***

### C. Normalize formats (Replace)

**Reformat date from `YYYY-MM-DD` → `DD/MM/YYYY`**

```vb
Set re = Rx("(\d{4})-(\d{2})-(\d{2})")
Debug.Print re.Replace("Due 2026-03-17", "Due $3/$2/$1")
```

**Standardize phone numbers (remove non-digits)**

```vb
Set re = Rx("[^\d]+", False, True)
Debug.Print re.Replace("(+39) 055-123 4567", "")
' → "390551234567"
```

**Collapse repeated whitespace to single space**

```vb
Set re = Rx("\s+", False, True)
Debug.Print re.Replace("alpha   beta   gamma", " ")
```

***

### D. Data cleaning

**Trim leading/trailing whitespace (including tabs/newlines)**

```vb
' Left trim
Debug.Print Rx("^\s+").Replace("  x  ", "")
' Right trim
Debug.Print Rx("\s+$").Replace("  x  ", "")
```

**Remove control characters (non‑printable)**

```vb
Debug.Print Rx("[\x00-\x1F\x7F]", False, True).Replace("A" & Chr(9) & "B", "")
```

**Strip HTML tags (simple, non-nested)**

```vb
Debug.Print Rx("<[^>]+>", False, True).Replace("<b>bold</b>", "bold")
```

***

### E. Parsing structured strings

**CSV line (basic, unquoted fields)**

```vb
Dim ms As Object, m As Object, fields() As String, i As Long
Set re = Rx("[^,]+", False, True)
Set ms = re.Execute("A,B,C,D")
ReDim fields(0 To ms.Count - 1)
i = 0: For Each m In ms: fields(i) = m.Value: i = i + 1: Next
```

**Key=Value pairs**

```vb
Set re = Rx("\b([A-Za-z_]\w*)=([^;\s]+)")
For Each m In re.Execute("user=alice id=42 role=admin")
    Debug.Print m.SubMatches(0), m.SubMatches(1)
Next
```

***

### F. Emails, URLs, and identifiers

**Email (pragmatic)**

```vb
Set re = Rx("\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", True)
```

**URL (http/https, simple)**

```vb
Set re = Rx("\bhttps?://[^\s)]+")
```

**IPv4**

```vb
Set re = Rx("\b(?:(?:25[0-5]|2[0-4]\d|1?\d?\d)\.){3}(?:25[0-5]|2[0-4]\d|1?\d?\d)\b")
```

**GUID**

```vb
Set re = Rx("\b[0-9A-F]{8}-([0-9A-F]{4}-){3}[0-9A-F]{12}\b")
```

***

### G. Numbers, currencies, and units

**Signed decimal (dot decimal, optional thousands sep)**

```vb
Set re = Rx("\b[+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?\b")
```

**Currency with symbol before amount**

```vb
Set re = Rx("\b(?:€|\$|£)\s?\d+(?:[\.,]\d{3})*(?:[.,]\d{2})?\b")
```

**Quantity with unit (kW, kg, m, mm, etc.)**

```vb
Set re = Rx("\b\d+(?:\.\d+)?\s?(?:kW|kg|m|mm|cm|km|°C|K|bar)\b", True)
```

***

### H. Lookarounds for context rules

**Find `A18370` only when preceded by `PO-` (lookbehind)**

```vb
Set re = Rx("(?<=PO-)[A-Z]\d{5}")
Debug.Print re.Test("PO-A18370") 'True
```

**Capture code not followed by a hyphen (negative lookahead)**

```vb
Set re = Rx("\b[A-Z]\d{5}\b(?!-)")
```

***

### I. Conditional extraction and masking

**Mask all but last 4 digits in numbers ≥ 8 digits**

```vb
Function MaskLongNumbers$(s$)
    Dim re As Object: Set re = Rx("(?<!\d)(\d{4,})(?!\d)", False, True)
    Dim m As Object
    For Each m In re.Execute(s)
        If Len(m.Value) >= 8 Then
            s = Replace(s, m.Value, String(Len(m.Value) - 4, "X") & Right(m.Value, 4))
        End If
    Next
    MaskLongNumbers = s
End Function
```

***

### J. Table‑like parsing (logs, semi‑structured text)

**Apache/Nginx common log (IP + request line + status)**

```vb
Set re = Rx("^(\S+) \S+ \S+ \[[^\]]+\] ""([A-Z]+) (.+?) HTTP/\d\.\d"" (\d{3})", False, True)
```

**INI section headers**

```vb
Set re = Rx("^\[([A-Za-z0-9_ -]+)\]$")
```

***

### K. Dates and times (multiple formats)

**Detect dd/mm/yyyy or dd-mm-yyyy**

```vb
Set re = Rx("\b(0?0?\d{4}\b")
```

**ISO 8601 date/time**

```vb
Set re = Rx("\b\d{4}-\d{2}-\d{2}(?:[ T]\d{2}:\d{2}:\d{2}(?:Z|[+-]\d{2}:\d{2})?)?\b")
```

***

### L. Excel-specific tasks

**Validate cell content with regex (UDF)**

```vb
Public Function RegexIsMatch(ByVal text As String, ByVal pattern As String, _
                             Optional ByVal ignoreCase As Boolean = False) As Boolean
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    With re
        .Pattern = pattern
        .IgnoreCase = ignoreCase
        .Global = False
    End With
    RegexIsMatch = re.Test(text)
End Function
' Usage in sheet:
' =RegexIsMatch(A2, "^[A-Z]\d{5}$")
```

**Extract first match (UDF)**

```vb
Public Function RegexExtract(ByVal text As String, ByVal pattern As String, _
                             Optional ByVal ignoreCase As Boolean = False) As String
    Dim re As Object, ms As Object
    Set re = CreateObject("VBScript.RegExp")
    With re
        .Pattern = pattern
        .IgnoreCase = ignoreCase
        .Global = False
    End With
    If re.Test(text) Then
        Set ms = re.Execute(text)
        RegexExtract = ms(0).Value
    Else
        RegexExtract = ""
    End If
End Function
' =RegexExtract(A2, "\b[A-Z]\d{5}\b")
```

**Extract nth capture group (UDF)**

```vb
Public Function RegexGroup(ByVal text As String, ByVal pattern As String, _
                           ByVal groupIndex As Long, _
                           Optional ByVal ignoreCase As Boolean = False) As String
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    With re
        .Pattern = pattern
        .IgnoreCase = ignoreCase
        .Global = False
    End With
    If re.Test(text) Then
        Set m = re.Execute(text)(0)
        If groupIndex >= 0 And groupIndex < m.SubMatches.Count Then
            RegexGroup = m.SubMatches(groupIndex)
        End If
    End If
End Function
' =RegexGroup(A2, "([A-Z])(\d{5})", 1)  -> "18370"
```

***

## Practical tips (VBScript.RegExp specifics)

*   Use **anchors** `^` and `$` to force full‑string validation; otherwise `.Test` returns true for any substring match.
*   Set `.Global = True` to iterate **all** matches (`.Execute`), otherwise only the first match is returned.
*   Replacement backreferences use **`$1`, `$2`, …** (not `\1`).
*   `.IgnoreCase = True` is the canonical way to enable case-insensitivity (inline flags like `(?i)` are not supported).
*   Named groups are **not supported**; use numeric groups (`SubMatches(0)` etc.).
*   For performance on large loops, **reuse the RegExp object** and avoid recreating it for each cell.

---

[DOC MOC](./doc-00_MOC.md)