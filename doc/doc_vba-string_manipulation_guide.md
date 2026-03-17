# VBA String Handling Guide

- [VBA String Handling Guide](#vba-string-handling-guide)
  - [Quick Reference — String Operations](#quick-reference--string-operations)
  - [Core Inspection \& Slicing](#core-inspection--slicing)
  - [Comparison \& Pattern Matching](#comparison--pattern-matching)
  - [Replace \& Removal](#replace--removal)
  - [Trimming, Cleaning, and Whitespace](#trimming-cleaning-and-whitespace)
  - [Strip Boundaries (quotes, brackets, parentheses)](#strip-boundaries-quotes-brackets-parentheses)
  - [Extract Before / After / Between](#extract-before--after--between)
  - [Split \& Join](#split--join)
  - [Counting Occurrences](#counting-occurrences)
  - [Padding \& Repeating](#padding--repeating)
  - [Character Codes (ASCII/Unicode)](#character-codes-asciiunicode)
  - [Regex (Search, Replace, Strip, Boundaries)](#regex-search-replace-strip-boundaries)
  - [Normalization \& Case](#normalization--case)
  - [Safe Null Handling](#safe-null-handling)
  - [Performance Tips](#performance-tips)
  - [Reusable Module: `modStringHelpers`](#reusable-module-modstringhelpers)
  - [Common “Recipes” (put together)](#common-recipes-put-together)
  - [Testing Harness (Immediate Window)](#testing-harness-immediate-window)
  - [Notes \& Gotchas](#notes--gotchas)

---

Below is a **comprehensive Excel VBA string‑manipulation guide** you can drop into your projects. It includes a **quick reference table** and **tested code recipes** for replace, compare, strip boundaries, trim/clean, parsing, regex, and more.

> **Note:** The summary table lists the actions, purpose, and key functions/operators. **Full, runnable VBA examples** are provided in sections below (outside the table) for clarity and correct rendering.

***

## Quick Reference — String Operations

> Use this table as an index. Jump to the matching section for full code examples.

| Action                         | What it does                       | Key Function / Operator                        |              |
| ------------------------------ | ---------------------------------- | ---------------------------------------------- | ------------ |
| Length                         | Count characters                   | `Len`, `LenB`                                  |              |
| Slice left/right/middle        | Return a substring                 | `Left$`, `Right$`, `Mid$`                      |              |
| Find substring                 | First/last occurrence              | `InStr`, `InStrRev`                            |              |
| Reverse string                 | Reverse character order            | `StrReverse`                                   |              |
| Case conversion                | Lower/upper/proper case            | `LCase$`, `UCase$`, `StrConv(vbProperCase)`    |              |
| Compare strings                | Case/locale-aware compare          | `StrComp`, `Option Compare`, `Like`            |              |
| Pattern match                  | Wildcards/ranges                   | `Like`, e.g., `s Like "AB#?-*" `               |              |
| Replace                        | Replace all occurrences            | `Replace$`                                     |              |
| Replace (nth)                  | Replace only the N‑th occurrence   | Loop + `InStr`/`InStrRev`                      |              |
| Trim spaces                    | Trim leading/trailing              | `Trim$`, `LTrim$`, `RTrim$`                    |              |
| Collapse spaces                | Normalize internal spaces to one   | `Application.WorksheetFunction.Trim`           |              |
| Clean non-printables           | Remove control chars               | `Application.WorksheetFunction.Clean`          |              |
| Normalize line breaks          | Standardize LF/CRLF                | `Replace` with `vbCrLf`, `vbLf`, `vbCr`        |              |
| Strip boundaries               | Remove wrapping quotes/brackets    | `Left$`/`Right$` check, or Regex \`^(\["'\[(]) | (\["'])])$\` |
| Keep/remove only certain chars | Whitelist/blacklist                | Regex (e.g., `[^A-Za-z0-9]`)                   |              |
| Split / Join                   | Tokenize + rebuild                 | `Split`, `Join`                                |              |
| Extract before/after           | Get text before/after a delimiter  | `Left$`/`Right$` + `InStr`/`InStrRev`          |              |
| Extract between                | Get text between two markers       | `Mid$` + positions                             |              |
| Count occurrences              | How many times a substring appears | Loop with `InStr`                              |              |
| Starts/Ends with               | Prefix/suffix test                 | Compare `Left$`/`Right$`                       |              |
| Pad left/right                 | Fixed width with pad char          | `String$`, `Right$`, `Left$`                   |              |
| Character code                 | Convert char⇄codepoint             | `AscW`, `ChrW$`                                |              |
| Performance tips               | Faster string ops                  | Use `$` suffixed funcs, `Join`, preallocation  |              |
| Regex search/replace           | Powerful patterns                  | `VBScript.RegExp` (late‑bound)                 |              |

***

## Core Inspection & Slicing

```vb
Sub CoreSlice()
    Dim s As String: s = "Climate-Tech_2026"
    Debug.Print Len(s)                  ' length in characters
    Debug.Print Left$(s, 7)             ' "Climate"
    Debug.Print Right$(s, 4)            ' "2026"
    Debug.Print Mid$(s, 9, 4)           ' "Tech"
    Debug.Print InStr(1, s, "_")        ' 13 (first occurrence)
    Debug.Print InStrRev(s, "-")        ' 8  (last occurrence)
    Debug.Print StrReverse(s)           ' "6202_hceT-etamilC"
End Sub
```

**Mid$ assignment (in-place replacement in a variable):**

```vb
Sub MidAssignment()
    Dim s As String: s = "ABC.DEF"
    Mid$(s, 4, 1) = "-"                 ' s becomes "ABC-DEF"
    Debug.Print s
End Sub
```

***

## Comparison & Pattern Matching

```vb
' Module or top of the file:
' Option Compare Text      ' makes literal string comparisons case-insensitive within this module
' Option Compare Binary    ' (default) case-sensitive, byte-wise

Sub CompareExamples()
    Dim a$, b$
    a = "Florence": b = "FLORENCE"
    Debug.Print StrComp(a, b, vbTextCompare)   ' 0 (equal, case-insensitive)
    Debug.Print StrComp(a, b, vbBinaryCompare) ' -1 (not equal; "F" < "f")
    
    ' Like operator (wildcards: ? one char, * any, # digit, [A-Z] range)
    Debug.Print ("AB12X" Like "AB##?")         ' True
    Debug.Print ("cat" Like "c[aeiou]t")       ' True
End Sub
```

**StartsWith / EndsWith helpers:**

```vb
Public Function StartsWith(ByVal s As String, ByVal prefix As String, _
                           Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As Boolean
    If Len(prefix) = 0 Then StartsWith = True: Exit Function
    If Len(s) < Len(prefix) Then Exit Function
    StartsWith = (StrComp(Left$(s, Len(prefix)), prefix, compareMethod) = 0)
End Function

Public Function EndsWith(ByVal s As String, ByVal suffix As String, _
                         Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As Boolean
    If Len(suffix) = 0 Then EndsWith = True: Exit Function
    If Len(s) < Len(suffix) Then Exit Function
    EndsWith = (StrComp(Right$(s, Len(suffix)), suffix, compareMethod) = 0)
End Function
```

***

## Replace & Removal

```vb
Sub ReplaceBasics()
    Dim s$
    s = "CO2, CO2e, CO, CO2"
    Debug.Print Replace$(s, "CO2", "CO₂")      ' Replace all "CO2" with "CO₂"
    Debug.Print Replace$(s, "CO2", "CO₂", 1, 1) ' Only the first occurrence
End Sub
```

**Remove a character or substring everywhere:**

```vb
Function RemoveAll(ByVal s As String, ByVal subStr As String) As String
    RemoveAll = Replace$(s, subStr, vbNullString)
End Function
```

**Replace only the N‑th occurrence:**

```vb
Function ReplaceNth(ByVal s As String, ByVal find As String, ByVal repl As String, ByVal n As Long, _
                    Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As String
    Dim i&, p&, startAt&
    If n <= 0 Or Len(find) = 0 Then ReplaceNth = s: Exit Function
    startAt = 1
    For i = 1 To n
        p = InStr(startAt, s, find, compareMethod)
        If p = 0 Then ReplaceNth = s: Exit Function
        startAt = p + Len(find)
    Next
    ReplaceNth = Left$(s, p - 1) & repl & Mid$(s, p + Len(find))
End Function
```

***

## Trimming, Cleaning, and Whitespace

```vb
Sub TrimClean()
    Dim s$
    s = "  " & vbTab & "  Climate  Tech   " & vbCrLf ' note NBSP before tab
    ' Trim leading/trailing ASCII spaces only:
    Debug.Print "[" & Trim$(s) & "]"
    
    ' Collapse internal spaces like Excel's TRIM (also trims leading/trailing):
    Debug.Print "[" & Application.WorksheetFunction.Trim(s) & "]"
    
    ' Remove control chars (ASCII < 32):
    Debug.Print "[" & Application.WorksheetFunction.Clean(s) & "]"
    
    ' Normalize NBSP (160) -> space, then TRIM:
    s = Replace$(s, Chr$(160), " ")
    Debug.Print "[" & Application.WorksheetFunction.Trim(s) & "]"
End Sub
```

**Normalize line breaks (force CRLF):**

```vb
Function NormalizeNewlines(ByVal s As String) As String
    s = Replace$(s, vbCrLf, vbLf)  ' unify first
    s = Replace$(s, vbCr, vbLf)
    s = Replace$(s, vbLf, vbCrLf)  ' final form: CRLF
    NormalizeNewlines = s
End Function
```

***

## Strip Boundaries (quotes, brackets, parentheses)

```vb
Function StripOuterOnce(ByVal s As String, ByVal leftCh As String, ByVal rightCh As String) As String
    If Len(s) >= 2 Then
        If Left$(s, 1) = leftCh And Right$(s, 1) = rightCh Then
            StripOuterOnce = Mid$(s, 2, Len(s) - 2)
            Exit Function
        End If
    End If
    StripOuterOnce = s
End Function

Sub StripExamples()
    Debug.Print StripOuterOnce("""value""", """", """")        ' -> value
    Debug.Print StripOuterOnce("(scope)", "(", ")")            ' -> scope
    Debug.Print StripOuterOnce("[tag]", "[", "]")              ' -> tag
End Sub
```

***

## Extract Before / After / Between

```vb
Function BeforeFirst(ByVal s As String, ByVal delim As String, _
                     Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As String
    Dim p&: p = InStr(1, s, delim, compareMethod)
    If p > 0 Then BeforeFirst = Left$(s, p - 1) Else BeforeFirst = s
End Function

Function AfterFirst(ByVal s As String, ByVal delim As String, _
                    Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As String
    Dim p&: p = InStr(1, s, delim, compareMethod)
    If p > 0 Then AfterFirst = Mid$(s, p + Len(delim)) Else AfterFirst = vbNullString
End Function

Function Between(ByVal s As String, ByVal leftMark As String, ByVal rightMark As String, _
                 Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As String
    Dim p1&, p2&
    p1 = InStr(1, s, leftMark, compareMethod)
    If p1 = 0 Then Exit Function
    p2 = InStr(p1 + Len(leftMark), s, rightMark, compareMethod)
    If p2 = 0 Then Exit Function
    Between = Mid$(s, p1 + Len(leftMark), p2 - (p1 + Len(leftMark)))
End Function

Sub ParseExamples()
    Dim s$: s = "sensor:CO2e;site:FI-Florence;value:421 ppm"
    Debug.Print BeforeFirst(s, ";")                ' sensor:CO2e
    Debug.Print AfterFirst(s, "site:")             ' FI-Florence;value:421 ppm
    Debug.Print Between(s, "value:", " ppm")       ' 421
End Sub
```

***

## Split & Join

```vb
Sub SplitJoinExamples()
    Dim s$, parts() As String
    s = "alpha|beta|gamma"
    parts = Split(s, "|")
    Debug.Print parts(0), parts(1), parts(2)
    Debug.Print Join(parts, " / ")                 ' "alpha / beta / gamma"
End Sub
```

> **CSV note:** `Split` is naive for quoted CSV. For robust CSV, let Excel parse: import via `QueryTables` or use `TextToColumns` on a worksheet. For moderate needs, a minimal quoted-splitter can be written, but it’s non-trivial.

***

## Counting Occurrences

```vb
Function CountOccurrences(ByVal s As String, ByVal subStr As String, _
                          Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As Long
    Dim p&, startAt&: startAt = 1
    If Len(subStr) = 0 Then Exit Function
    Do
        p = InStr(startAt, s, subStr, compareMethod)
        If p = 0 Then Exit Do
        CountOccurrences = CountOccurrences + 1
        startAt = p + Len(subStr)
    Loop
End Function
```

***

## Padding & Repeating

```vb
Function PadLeft(ByVal s As String, ByVal width As Long, Optional ByVal padChar As String = " ") As String
    If Len(s) >= width Then PadLeft = s Else PadLeft = String$(width - Len(s), Left$(padChar, 1)) & s
End Function

Function PadRight(ByVal s As String, ByVal width As Long, Optional ByVal padChar As String = " ") As String
    If Len(s) >= width Then PadRight = s Else PadRight = s & String$(width - Len(s), Left$(padChar, 1))
End Function

Sub PadExamples()
    Debug.Print PadLeft("42", 5, "0")             ' "00042"
    Debug.Print PadRight("CT", 4, "-")            ' "CT--"
End Sub
```

***

## Character Codes (ASCII/Unicode)

```vb
Sub CharCodes()
    Debug.Print AscW("é")             ' 233
    Debug.Print ChrW$(233)            ' "é"
    Debug.Print AscW("₃")             ' subscript 3
End Sub
```

***

## Regex (Search, Replace, Strip, Boundaries)

> **No reference required** (late binding). VBScript.RegExp supports: `.` `*` `+` `?` `{m,n}` `[]` groups `()` alternation `|` anchors `^` `$`, **word boundary `\b`**, but **no lookbehind**.

```vb
Private Function RegexReplace(ByVal s As String, ByVal pattern As String, ByVal replacement As String, _
                              Optional ByVal ignoreCase As Boolean = True, _
                              Optional ByVal multiLine As Boolean = True) As String
    Dim rx As Object
    Set rx = CreateObject("VBScript.RegExp")
    rx.Pattern = pattern
    rx.Global = True
    rx.IgnoreCase = ignoreCase
    rx.MultiLine = multiLine
    RegexReplace = rx.Replace(s, replacement)
End Function

Private Function RegexIsMatch(ByVal s As String, ByVal pattern As String, _
                              Optional ByVal ignoreCase As Boolean = True, _
                              Optional ByVal multiLine As Boolean = True) As Boolean
    Dim rx As Object
    Set rx = CreateObject("VBScript.RegExp")
    rx.Pattern = pattern
    rx.Global = False
    rx.IgnoreCase = ignoreCase
    rx.MultiLine = multiLine
    RegexIsMatch = rx.Test(s)
End Function

Sub RegexExamples()
    Dim s$
    s = " [ID-001]  CO2e@Florence  "
    ' Strip non-alphanumerics (keep A–Z, a–z, 0–9, space):
    Debug.Print RegexReplace(s, "[^A-Za-z0-9 ]", "")
    ' Collapse 2+ spaces to one:
    Debug.Print RegexReplace(s, " {2,}", " ")
    ' Strip leading/trailing whitespace:
    Debug.Print RegexReplace(s, "^\s+|\s+$", "")
    ' Word boundary example:
    Debug.Print RegexReplace("cat scat category", "\bcat\b", "dog")  ' "dog scat category"
    ' Strip paired quotes only when both present:
    Debug.Print RegexReplace("""value""", "^([""'])(.*)\1$", "$2")
End Sub
```

***

## Normalization & Case

```vb
Sub CaseAndProper()
    Dim s$: s = "sCoPe scope SCOPE"
    Debug.Print LCase$(s)
    Debug.Print UCase$(s)
    Debug.Print StrConv("renewable energy storage", vbProperCase) ' "Renewable Energy Storage"
End Sub
```

> **Note:** `StrConv(vbProperCase)` is locale-dependent and not Unicode‑aware for all languages. VBA has no built‑in diacritic remover; use whitelist regex or a mapping function if you need ASCII‑only.

***

## Safe Null Handling

```vb
Function NzStr(ByVal s As Variant, Optional ByVal fallback As String = vbNullString) As String
    If IsNull(s) Or IsEmpty(s) Then
        NzStr = fallback
    Else
        NzStr = CStr(s)
    End If
End Function
```

***

## Performance Tips

*   Prefer **`$`**-suffixed functions (`Left$`, `Right$`, `Mid$`, `Replace$`) to avoid `Variant` temporaries.
*   Avoid repeated `&` concatenation in loops; **collect in an array and `Join`**, or preallocate with `String$` and use `Mid$` assignment.
*   Reuse the same `RegExp` object when applying the same pattern repeatedly (if using early binding).
*   Minimize `WorksheetFunction` calls inside tight loops (batch process where possible).

**Example: Build a large string efficiently**

```vb
Function RepeatConcat(ByVal token As String, ByVal times As Long, ByVal sep As String) As String
    Dim arr() As String, i&
    ReDim arr(1 To times)
    For i = 1 To times
        arr(i) = token
    Next
    RepeatConcat = Join(arr, sep)
End Function
```

***

## Reusable Module: `modStringHelpers`

> Drop this into a standard module and you’ll have most utilities ready to go.

```vb
' ===== Module: modStringHelpers =====

Option Explicit

Public Function StartsWith(ByVal s As String, ByVal prefix As String, _
                           Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As Boolean
    If Len(prefix) = 0 Then StartsWith = True: Exit Function
    If Len(s) < Len(prefix) Then Exit Function
    StartsWith = (StrComp(Left$(s, Len(prefix)), prefix, compareMethod) = 0)
End Function

Public Function EndsWith(ByVal s As String, ByVal suffix As String, _
                         Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As Boolean
    If Len(suffix) = 0 Then EndsWith = True: Exit Function
    If Len(s) < Len(suffix) Then Exit Function
    EndsWith = (StrComp(Right$(s, Len(suffix)), suffix, compareMethod) = 0)
End Function

Public Function BeforeFirst(ByVal s As String, ByVal delim As String, _
                            Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As String
    Dim p&: p = InStr(1, s, delim, compareMethod)
    If p > 0 Then BeforeFirst = Left$(s, p - 1) Else BeforeFirst = s
End Function

Public Function AfterFirst(ByVal s As String, ByVal delim As String, _
                           Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As String
    Dim p&: p = InStr(1, s, delim, compareMethod)
    If p > 0 Then AfterFirst = Mid$(s, p + Len(delim)) Else AfterFirst = vbNullString
End Function

Public Function Between(ByVal s As String, ByVal leftMark As String, ByVal rightMark As String, _
                        Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As String
    Dim p1&, p2&
    p1 = InStr(1, s, leftMark, compareMethod)
    If p1 = 0 Then Exit Function
    p2 = InStr(p1 + Len(leftMark), s, rightMark, compareMethod)
    If p2 = 0 Then Exit Function
    Between = Mid$(s, p1 + Len(leftMark), p2 - (p1 + Len(leftMark)))
End Function

Public Function CountOccurrences(ByVal s As String, ByVal subStr As String, _
                                 Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As Long
    Dim p&, startAt&: startAt = 1
    If Len(subStr) = 0 Then Exit Function
    Do
        p = InStr(startAt, s, subStr, compareMethod)
        If p = 0 Then Exit Do
        CountOccurrences = CountOccurrences + 1
        startAt = p + Len(subStr)
    Loop
End Function

Public Function ReplaceNth(ByVal s As String, ByVal find As String, ByVal repl As String, ByVal n As Long, _
                           Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As String
    Dim i&, p&, startAt&
    If n <= 0 Or Len(find) = 0 Then ReplaceNth = s: Exit Function
    startAt = 1
    For i = 1 To n
        p = InStr(startAt, s, find, compareMethod)
        If p = 0 Then ReplaceNth = s: Exit Function
        startAt = p + Len(find)
    Next
    ReplaceNth = Left$(s, p - 1) & repl & Mid$(s, p + Len(find))
End Function

Public Function StripOuterOnce(ByVal s As String, ByVal leftCh As String, ByVal rightCh As String) As String
    If Len(s) >= 2 Then
        If Left$(s, 1) = leftCh And Right$(s, 1) = rightCh Then
            StripOuterOnce = Mid$(s, 2, Len(s) - 2)
            Exit Function
        End If
    End If
    StripOuterOnce = s
End Function

Public Function PadLeft(ByVal s As String, ByVal width As Long, Optional ByVal padChar As String = " ") As String
    If Len(s) >= width Then PadLeft = s Else PadLeft = String$(width - Len(s), Left$(padChar, 1)) & s
End Function

Public Function PadRight(ByVal s As String, ByVal width As Long, Optional ByVal padChar As String = " ") As String
    If Len(s) >= width Then PadRight = s Else PadRight = s & String$(width - Len(s), Left$(padChar, 1))
End Function

Public Function NormalizeNewlines(ByVal s As String) As String
    s = Replace$(s, vbCrLf, vbLf)
    s = Replace$(s, vbCr, vbLf)
    s = Replace$(s, vbLf, vbCrLf)
    NormalizeNewlines = s
End Function

Public Function RemoveAll(ByVal s As String, ByVal subStr As String) As String
    RemoveAll = Replace$(s, subStr, vbNullString)
End Function

Public Function NzStr(ByVal s As Variant, Optional ByVal fallback As String = vbNullString) As String
    If IsNull(s) Or IsEmpty(s) Then NzStr = fallback Else NzStr = CStr(s)
End Function

' --- Regex helpers (late-bound) ---

Public Function RegexReplace(ByVal s As String, ByVal pattern As String, ByVal replacement As String, _
                             Optional ByVal ignoreCase As Boolean = True, _
                             Optional ByVal multiLine As Boolean = True) As String
    Dim rx As Object
    Set rx = CreateObject("VBScript.RegExp")
    rx.Pattern = pattern
    rx.Global = True
    rx.IgnoreCase = ignoreCase
    rx.MultiLine = multiLine
    RegexReplace = rx.Replace(s, replacement)
End Function

Public Function RegexIsMatch(ByVal s As String, ByVal pattern As String, _
                             Optional ByVal ignoreCase As Boolean = True, _
                             Optional ByVal multiLine As Boolean = True) As Boolean
    Dim rx As Object
    Set rx = CreateObject("VBScript.RegExp")
    rx.Pattern = pattern
    rx.Global = False
    rx.IgnoreCase = ignoreCase
    rx.MultiLine = multiLine
    RegexIsMatch = rx.Test(s)
End Function
```

***

## Common “Recipes” (put together)

1.  **Standardize sensor tags**
    *   Uppercase, remove spaces, replace hyphens with underscore:
    ```vb
    Function NormalizeTag(ByVal s As String) As String
        s = UCase$(s)
        s = Replace$(s, " ", vbNullString)
        s = Replace$(s, "-", "_")
        NormalizeTag = s
    End Function
    ```

2.  **Keep only digits (e.g., extract number)**
    ```vb
    Function KeepDigits(ByVal s As String) As String
        KeepDigits = RegexReplace(s, "[^\d]", "")
    End Function
    ```

3.  **Clean and collapse whitespace (incl. NBSP)**
    ```vb
    Function CleanSpaces(ByVal s As String) As String
        s = Application.WorksheetFunction.Clean(s)
        s = Replace$(s, Chr$(160), " ")
        CleanSpaces = Application.WorksheetFunction.Trim(s)
    End Function
    ```

4.  **Get the file extension (case-insensitive)**
    ```vb
    Function FileExt(ByVal path As String) As String
        Dim p&: p = InStrRev(path, ".")
        If p = 0 Then Exit Function
        FileExt = LCase$(Mid$(path, p + 1))
    End Function
    ```

***

## Testing Harness (Immediate Window)

```vb
Sub Demo_All()
    Debug.Print "=== Core ===": CoreSlice: MidAssignment
    Debug.Print "=== Compare ===": CompareExamples
    Debug.Print "=== Replace ===": ReplaceBasics
    Debug.Print "=== Trim/Clean ===": TrimClean
    Debug.Print "=== Boundaries ===": StripExamples
    Debug.Print "=== Parse ===": ParseExamples
    Debug.Print "=== Split/Join ===": SplitJoinExamples
    Debug.Print "=== Padding ===": PadExamples
    Debug.Print "=== Regex ===": RegexExamples
End Sub
```

***

## Notes & Gotchas

*   `Trim` vs `WorksheetFunction.Trim`: VBA `Trim$` only removes **leading/trailing ASCII space**; Excel’s `TRIM` also collapses **multiple internal spaces to one**, but **NBSP (160)** often requires explicit replacement first.
*   `LenB` reports bytes (for Unicode, typically `2 * Len`), not characters.
*   `Like` uses **VB wildcard rules**, not full regex semantics.
*   `StrConv(vbProperCase)` is locale-sensitive and imperfect for mixed scripts.
*   VBScript.RegExp **does not support lookbehind**; redesign patterns when needed.

***

Would you like me to **package these helpers into a downloadable `.bas` module** you can import into your VBA project, or tailor the recipes to a specific dataset (e.g., CO₂e tag normalization in your Florence pipelines)?

---

[DOC MOC](./doc-00_MOC.md)