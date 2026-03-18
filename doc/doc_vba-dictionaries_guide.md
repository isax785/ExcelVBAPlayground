# VBA Dictionaries

- [VBA Dictionaries](#vba-dictionaries)
  - [1) What is a `Dictionary` and when to use it](#1-what-is-a-dictionary-and-when-to-use-it)
  - [2) Early vs Late Binding](#2-early-vs-late-binding)
  - [3) Dictionary API — Main Operations (Summary)](#3-dictionary-api--main-operations-summary)
  - [4) Best Practices \& Pitfalls (Excel-Focused)](#4-best-practices--pitfalls-excel-focused)
  - [5) Reusable Building Blocks (Copy‑Ready)](#5-reusable-building-blocks-copyready)
  - [6) Building Dictionaries from Ranges — Core Patterns](#6-building-dictionaries-from-ranges--core-patterns)
    - [Example A — Two‑Column Lookup (Key → Value) from a Range](#example-a--twocolumn-lookup-key--value-from-a-range)
    - [Example B — Frequency Count / Unique Values from a Single Column](#example-b--frequency-count--unique-values-from-a-single-column)
    - [Example C — Grouping: Key → Collection of Values (handle duplicates)](#example-c--grouping-key--collection-of-values-handle-duplicates)
    - [Example D — Dictionary-Powered VLOOKUP (left join)](#example-d--dictionary-powered-vlookup-left-join)
    - [Example E — Merge Two Dictionaries with Conflict Resolution](#example-e--merge-two-dictionaries-with-conflict-resolution)
    - [Example F — Sort by Keys (or Values) for Reporting](#example-f--sort-by-keys-or-values-for-reporting)
    - [Example G — Header Map: Header → Column Index](#example-g--header-map-header--column-index)
    - [Example H — Composite Keys from Multi-Column Ranges](#example-h--composite-keys-from-multi-column-ranges)
    - [Example I — De‑duplicate a Range (Preserve First Occurrence Order)](#example-i--deduplicate-a-range-preserve-first-occurrence-order)
    - [Example J — Safe Removal While Iterating](#example-j--safe-removal-while-iterating)
  - [7) Performance Checklist (Excel)](#7-performance-checklist-excel)
  - [8) Troubleshooting \& Edge Cases](#8-troubleshooting--edge-cases)
  - [9) Quick “Recipes” Index](#9-quick-recipes-index)
  - [10) Optional: Early Binding Variants](#10-optional-early-binding-variants)
- [Error Handling](#error-handling)
  - [1) Error Handling Patterns (Quick Summary)](#1-error-handling-patterns-quick-summary)
  - [2) Common Dictionary-Related Errors (Reference)](#2-common-dictionary-related-errors-reference)
  - [3) Guarded, Structured Error Handling Templates](#3-guarded-structured-error-handling-templates)
    - [3.1 Validate a Range (Guard Clause)](#31-validate-a-range-guard-clause)
    - [3.2 Centralized Error Reporter (Reusable)](#32-centralized-error-reporter-reusable)
    - [3.3 Structured Handler with Cleanup](#33-structured-handler-with-cleanup)
  - [4) “Try\*” Functions (Boolean + ByRef) — No Exceptions Bubble Up](#4-try-functions-boolean--byref--no-exceptions-bubble-up)
    - [4.1 Try to Build Dictionary from Two Columns](#41-try-to-build-dictionary-from-two-columns)
    - [4.2 Try Fill (Dictionary-Powered VLOOKUP), Robust to Errors in Source Range](#42-try-fill-dictionary-powered-vlookup-robust-to-errors-in-source-range)
  - [5) Safe Wrappers for Specific Dictionary Operations](#5-safe-wrappers-for-specific-dictionary-operations)
    - [5.1 Safe Add (Avoid Err 457)](#51-safe-add-avoid-err-457)
    - [5.2 Safe Upsert](#52-safe-upsert)
    - [5.3 Safe Merge with Conflict Policy (Prefer Left/Right)](#53-safe-merge-with-conflict-policy-prefer-leftright)
    - [5.4 Safe Rename Key (Prevent Collision)](#54-safe-rename-key-prevent-collision)
  - [6) Robust Range → Dictionary Builders (With Error Handling)](#6-robust-range--dictionary-builders-with-error-handling)
    - [6.1 Two-Column Builder with Explicit Error Path](#61-two-column-builder-with-explicit-error-path)
    - [6.2 Composite Key Builder with Defensive Checks](#62-composite-key-builder-with-defensive-checks)
  - [7) Handling Cell Errors and Blanks Explicitly](#7-handling-cell-errors-and-blanks-explicitly)
    - [7.1 Skip or Map Excel Errors](#71-skip-or-map-excel-errors)
  - [8) Defensive Iteration and Removal](#8-defensive-iteration-and-removal)
    - [8.1 Remove by Predicate with Full Safety](#81-remove-by-predicate-with-full-safety)
  - [9) Logging \& Escalation Strategy](#9-logging--escalation-strategy)
  - [10) End-to-End Example: Safe De‑duplication Macro](#10-end-to-end-example-safe-deduplication-macro)

---

Below is a practical, end‑to‑end guide to `Scripting.Dictionary` in VBA, with a strong focus on building dictionaries from Excel ranges. I’ve included:

*   A concise **markdown table** summarizing the main operations (no long code in the table; full code comes in the examples).
*   A comprehensive set of **reusable procedures and functions** for real‑world tasks.
*   Best practices, pitfalls, and performance tips specifically for Excel/VBA.

> **Applies to**: Excel VBA (Windows/Mac).  
> **Reference model**: `Scripting.Dictionary` from **Microsoft Scripting Runtime** (`scrrun.dll`). You can use **early binding** (set reference) or **late binding** (`CreateObject`).

***

## 1) What is a `Dictionary` and when to use it

A `Dictionary` maps **unique keys** to **values**, providing fast insert/lookup/update. Use it when you need:

*   **De‑duplication / unique lists**
*   **Lookups** (as a faster and safer alternative to repeated `VLOOKUP`)
*   **Counting/frequency** of items
*   **Grouping** (key → many values)
*   **Join-like operations** between ranges/tables
*   **Indexing headers** (header → column index)

Compared to `Collection`:

*   `Dictionary.Exists(key)` is O(1) and avoids error‑prone lookups.
*   You can choose text vs binary comparisons (`CompareMode`).
*   You can read all keys/values via `Keys` / `Items`.

***

## 2) Early vs Late Binding

*   **Early binding** (recommended for dev): Add reference **Tools → References → Microsoft Scripting Runtime**; then `Dim dict As Scripting.Dictionary`.
*   **Late binding** (deployment w/o references): `Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")`.

> **Tip**: You can write helper factory functions to centralize the choice.

***

## 3) Dictionary API — Main Operations (Summary)

> **Note**: The “Usage” column shows short one‑liners (not full code). Complete, copy‑ready code is provided in the examples below.

| Operation        | Purpose                             | Usage                                                                                                     | Notes                                                                 |
| ---------------- | ----------------------------------- | --------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------- |
| Create           | Instantiate a new dictionary        | `Set dict = New Scripting.Dictionary` (early) or `Set dict = CreateObject("Scripting.Dictionary")` (late) | Prefer a factory that sets `CompareMode` up front.                    |
| CompareMode      | Case sensitivity for string keys    | `dict.CompareMode = vbTextCompare`                                                                        | **Set this before adding items.** Binary is default (case-sensitive). |
| Add              | Add key/value (error if key exists) | `dict.Add key, value`                                                                                     | Use when you want to enforce uniqueness strictly.                     |
| Assign/Update    | Set or overwrite value              | `dict(key) = value`                                                                                       | Creates if missing; updates if existing.                              |
| Exists           | Test if key exists                  | `If dict.Exists(key) Then ...`                                                                            | Use before `Add` to avoid errors.                                     |
| Item             | Get or set value                    | `value = dict(key)` or `dict(key) = newValue`                                                             | Raises error if getting a missing key (check `Exists` first).         |
| Remove           | Remove one                          | `dict.Remove key`                                                                                         | Safe only if you’re sure it exists (or check `Exists`).               |
| RemoveAll        | Clear all                           | `dict.RemoveAll`                                                                                          | Frees entries but keeps the object.                                   |
| Count            | Size                                | `n = dict.Count`                                                                                          | Useful for sizing arrays before writing back to sheets.               |
| Keys             | All keys as `Variant()`             | `arrK = dict.Keys`                                                                                        | Unsigned order; returns a snapshot array.                             |
| Items            | All values as `Variant()`           | `arrV = dict.Items`                                                                                       | Often paired with `Keys`.                                             |
| For Each         | Iterate keys                        | `For Each k In dict.Keys ... Next`                                                                        | Don’t mutate while iterating; collect keys first if removing.         |
| Rename key       | Change key                          | `dict.Key(oldKey) = newKey`                                                                               | Works, but **use sparingly**; prefer remove/add.                      |
| Default property | Bracket access                      | `dict(key)`                                                                                               | Useful for “upsert”-style writes.                                     |

***

## 4) Best Practices & Pitfalls (Excel-Focused)

*   **Set `CompareMode` before adding any items** (`vbTextCompare` for case‑insensitive text).
*   **Normalize keys** (e.g., `Trim`, `LCase`) from ranges to avoid unseen mismatches.
*   **Load ranges into arrays** first; iterate arrays, not cells (huge performance win).
*   **Avoid modifying the dictionary while iterating**; collect keys to remove first.
*   **`Keys` and `Items` order is not guaranteed**; to sort, export to arrays and sort there.
*   **Handle blanks and errors** in ranges explicitly.
*   **Composite keys**: join fields with a delimiter **that cannot appear in data** (e.g., `ChrW(30)`).

***

## 5) Reusable Building Blocks (Copy‑Ready)

> Place these in a standard VBA module. All examples work with **early or late binding**. If you can’t set a reference, switch `As Scripting.Dictionary` to `As Object` and use `CreateObject`.

```vb
Option Explicit

'=== Factory: create Dictionary with chosen CompareMode ===
Public Function NewDict(Optional ByVal caseInsensitive As Boolean = True) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary") 'works w/ or w/o reference
    d.CompareMode = IIf(caseInsensitive, vbTextCompare, vbBinaryCompare)
    Set NewDict = d
End Function

'=== Normalize a key safely (avoid Null/errors) ===
Public Function NormalizeKey(ByVal v As Variant) As String
    On Error Resume Next
    NormalizeKey = LCase$(Trim$(CStr(v)))
End Function

'=== Read a Range to a 2D Variant array (fast) ===
Public Function RngToArray(ByVal rng As Range) As Variant
    If rng Is Nothing Then
        RngToArray = Empty
        Exit Function
    End If
    If rng.Rows.Count = 1 And rng.Columns.Count = 1 Then
        ' Wrap single cell into 2D [1..1,1..1]
        Dim arr(1 To 1, 1 To 1) As Variant
        arr(1, 1) = rng.Value
        RngToArray = arr
    Else
        RngToArray = rng.Value2 'Value2 avoids currency/date conversions
    End If
End Function

'=== Safely write a 2D array back to a Range ===
Public Sub ArrayToRange(ByVal target As Range, ByRef arr As Variant)
    If IsEmpty(arr) Then Exit Sub
    Dim r As Long, c As Long
    r = UBound(arr, 1) - LBound(arr, 1) + 1
    c = UBound(arr, 2) - LBound(arr, 2) + 1
    target.Resize(r, c).Value = arr
End Sub

'=== Make a composite key from multiple parts (collision-resistant) ===
Public Function MakeKey(ParamArray parts() As Variant) As String
    Const SEP As String = vbNullChar & ChrW(30) & vbNullChar 'unlikely in real data
    Dim i As Long, tmp() As String
    ReDim tmp(LBound(parts) To UBound(parts))
    For i = LBound(parts) To UBound(parts)
        tmp(i) = NormalizeKey(parts(i))
    Next
    MakeKey = Join(tmp, SEP)
End Function

'=== Simple in-place QuickSort for paired arrays (sort by keys; keep values aligned) ===
Public Sub QuickSortKV(ByRef k As Variant, ByRef v As Variant, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long
    Dim pivot As Variant
    Dim tmpK As Variant, tmpV As Variant
    i = lo: j = hi
    pivot = k((lo + hi) \ 2)
    Do While i <= j
        Do While k(i) < pivot: i = i + 1: Loop
        Do While k(j) > pivot: j = j - 1: Loop
        If i <= j Then
            tmpK = k(i): k(i) = k(j): k(j) = tmpK
            tmpV = v(i): v(i) = v(j): v(j) = tmpV
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then QuickSortKV k, v, lo, j
    If i < hi Then QuickSortKV k, v, i, hi
End Sub
```

***

## 6) Building Dictionaries from Ranges — Core Patterns

### Example A — Two‑Column Lookup (Key → Value) from a Range

**Task**: Build a dictionary from a two‑column range (e.g., `A2:B100`), ignoring blanks and trimming/normalizing keys.

```vb
Public Function DictFromTwoColumns(ByVal rng As Range, _
                                   Optional ByVal caseInsensitive As Boolean = True) As Object
    Dim d As Object: Set d = NewDict(caseInsensitive)
    Dim arr As Variant, r As Long, key As String, val As Variant
    arr = RngToArray(rng)
    If IsEmpty(arr) Then Set DictFromTwoColumns = d: Exit Function
    
    For r = LBound(arr, 1) To UBound(arr, 1)
        key = NormalizeKey(arr(r, 1))
        val = arr(r, 2)
        If Len(key) > 0 Then
            ' Prefer first mapping; skip duplicates unless you want last-in-wins:
            If Not d.Exists(key) Then d.Add key, val
            ' Or: d(key) = val 'last-in-wins
        End If
    Next
    Set DictFromTwoColumns = d
End Function
```

**Use**:

```vb
Sub BuildAndUseLookup()
    Dim d As Object
    Set d = DictFromTwoColumns(Worksheets("Data").Range("A2:B100"))
    
    If d.Exists("code123") Then
        Debug.Print "Value for code123:", d("code123")
    End If
End Sub
```

***

### Example B — Frequency Count / Unique Values from a Single Column

**Task**: Count occurrences in a single column (e.g., `A2:A1000`), write unique + counts to output.

```vb
Public Sub FrequencyFromColumn()
    Dim src As Range, out As Range, arr As Variant
    Dim d As Object, r As Long, key As String
    Set src = Worksheets("Data").Range("A2").CurrentRegion.Columns(1) 'or a defined range
    arr = RngToArray(src)
    Set d = NewDict(True)
    
    For r = LBound(arr, 1) To UBound(arr, 1)
        key = NormalizeKey(arr(r, 1))
        If Len(key) > 0 Then
            d(key) = IIf(d.Exists(key), d(key) + 1, 1)
        End If
    Next
    
    ' Write out
    Dim k As Variant, v As Variant, i As Long, n As Long
    k = d.Keys: v = d.Items
    n = d.Count
    If n = 0 Then Exit Sub
    
    Dim outArr() As Variant
    ReDim outArr(1 To n, 1 To 2)
    For i = 1 To n
        outArr(i, 1) = k(i - 1)
        outArr(i, 2) = v(i - 1)
    Next
    
    Set out = Worksheets("Report").Range("D2")
    ArrayToRange out, outArr
End Sub
```

***

### Example C — Grouping: Key → Collection of Values (handle duplicates)

**Task**: Build groups where each key maps to multiple values (e.g., product → all order IDs).

```vb
Public Function DictGroup(ByVal rng As Range) As Object
    ' rng expects two columns: [Key][Value]
    Dim d As Object: Set d = NewDict(True)
    Dim arr As Variant, r As Long, key As String, val As Variant
    arr = RngToArray(rng)
    If IsEmpty(arr) Then Set DictGroup = d: Exit Function
    
    For r = LBound(arr, 1) To UBound(arr, 1)
        key = NormalizeKey(arr(r, 1))
        val = arr(r, 2)
        If Len(key) > 0 Then
            If Not d.Exists(key) Then Set d(key) = New Collection
            d(key).Add val
        End If
    Next
    Set DictGroup = d
End Function
```

***

### Example D — Dictionary-Powered VLOOKUP (left join)

**Task**: Replace `VLOOKUP`: map codes→prices, then populate prices for a list of codes efficiently.

```vb
Public Sub FillPricesWithDict()
    Dim mapRng As Range, codesRng As Range, outRng As Range
    Set mapRng = Worksheets("Data").Range("A2:B100") ' [Code][Price]
    Set codesRng = Worksheets("Data").Range("D2:D500") ' codes to lookup
    Set outRng = codesRng.Offset(0, 1) ' write prices in column E
    
    Dim d As Object: Set d = DictFromTwoColumns(mapRng, True)
    Dim arrIn As Variant, arrOut() As Variant, i As Long, n As Long, key As String
    arrIn = RngToArray(codesRng)
    n = UBound(arrIn, 1)
    ReDim arrOut(1 To n, 1 To 1)
    
    For i = 1 To n
        key = NormalizeKey(arrIn(i, 1))
        If d.Exists(key) Then
            arrOut(i, 1) = d(key)
        Else
            arrOut(i, 1) = CVErr(xlErrNA) ' or leave blank
        End If
    Next
    ArrayToRange outRng, arrOut
End Sub
```

***

### Example E — Merge Two Dictionaries with Conflict Resolution

**Task**: Merge `left` and `right` dictionaries; decide which value wins on key conflicts.

```vb
Public Function MergeDicts(ByVal left As Object, ByVal right As Object, _
                           Optional ByVal preferRight As Boolean = True) As Object
    Dim d As Object: Set d = NewDict(True)
    Dim k As Variant, i As Long
    
    'copy left
    k = left.Keys
    For i = LBound(k) To UBound(k)
        d(k(i)) = left(k(i))
    Next
    
    'merge right
    k = right.Keys
    For i = LBound(k) To UBound(k)
        If d.Exists(k(i)) Then
            If preferRight Then d(k(i)) = right(k(i))
        Else
            d(k(i)) = right(k(i))
        End If
    Next
    Set MergeDicts = d
End Function
```

***

### Example F — Sort by Keys (or Values) for Reporting

**Task**: Dump keys & values into arrays and sort them (dictionary itself is unsorted).

```vb
Public Sub ReportSortedKeys(ByVal d As Object, ByVal target As Range)
    If d Is Nothing Then Exit Sub
    If d.Count = 0 Then Exit Sub
    
    Dim k As Variant, v As Variant, n As Long, i As Long, out() As Variant
    k = d.Keys: v = d.Items: n = d.Count
    ' Convert to 1-based arrays for QuickSort
    Dim keys() As Variant, vals() As Variant
    ReDim keys(1 To n): ReDim vals(1 To n)
    For i = 1 To n
        keys(i) = k(i - 1)
        vals(i) = v(i - 1)
    Next
    
    QuickSortKV keys, vals, 1, n
    
    ReDim out(1 To n, 1 To 2)
    For i = 1 To n
        out(i, 1) = keys(i)
        out(i, 2) = vals(i)
    Next
    ArrayToRange target, out
End Sub
```

***

### Example G — Header Map: Header → Column Index

**Task**: Build a dictionary for dynamic column addressing.

```vb
Public Function HeaderIndexMap(ByVal headerRow As Range) As Object
    Dim d As Object: Set d = NewDict(True)
    Dim arr As Variant, c As Long, key As String
    arr = RngToArray(headerRow)
    For c = LBound(arr, 2) To UBound(arr, 2)
        key = NormalizeKey(arr(1, c))
        If Len(key) > 0 And Not d.Exists(key) Then d.Add key, c
    Next
    Set HeaderIndexMap = d
End Function
```

**Use**:

```vb
Sub DemoHeaderIndex()
    Dim d As Object: Set d = HeaderIndexMap(Worksheets("Data").Range("A1:Z1"))
    Dim col As Long
    If d.Exists("price") Then
        col = d("price") ' 1-based column within A1:Z1
        Debug.Print "Price column offset in the header row:", col
    End If
End Sub
```

***

### Example H — Composite Keys from Multi-Column Ranges

**Task**: Treat multiple columns as a unique identifier (e.g., `[Region, Product, Month]`).

```vb
Public Function DictFromComposite(ByVal rng As Range) As Object
    ' Range with >= 2 columns; key = first n-1 columns; value = last column
    Dim d As Object: Set d = NewDict(True)
    Dim arr As Variant, r As Long, c As Long, lastCol As Long
    arr = RngToArray(rng)
    If IsEmpty(arr) Then Set DictFromComposite = d: Exit Function
    
    lastCol = UBound(arr, 2)
    For r = LBound(arr, 1) To UBound(arr, 1)
        Dim parts() As Variant
        ReDim parts(1 To lastCol - 1)
        For c = 1 To lastCol - 1
            parts(c) = arr(r, c)
        Next
        Dim key As String: key = MakeKey(parts)
        If Len(key) > 0 Then
            d(key) = arr(r, lastCol)
        End If
    Next
    Set DictFromComposite = d
End Function
```

***

### Example I — De‑duplicate a Range (Preserve First Occurrence Order)

**Task**: Produce a list of unique items in the order they first appear.

```vb
Public Sub UniquePreserveOrder()
    Dim src As Range, arr As Variant, d As Object
    Set src = Worksheets("Data").Range("A2:A1000")
    arr = RngToArray(src)
    Set d = NewDict(True)
    
    Dim r As Long, key As String
    For r = LBound(arr, 1) To UBound(arr, 1)
        key = NormalizeKey(arr(r, 1))
        If Len(key) > 0 Then
            If Not d.Exists(key) Then d.Add key, arr(r, 1) 'store original value for output
        End If
    Next
    
    Dim k As Variant, v As Variant, n As Long, i As Long, out() As Variant
    n = d.Count: If n = 0 Then Exit Sub
    k = d.Keys: v = d.Items
    ReDim out(1 To n, 1 To 1)
    For i = 1 To n
        out(i, 1) = v(i - 1)
    Next
    ArrayToRange Worksheets("Report").Range("B2"), out
End Sub
```

***

### Example J — Safe Removal While Iterating

**Task**: Remove entries by predicate without mutating during `For Each`.

```vb
Public Sub RemoveShortKeys(ByVal d As Object, ByVal minLen As Long)
    Dim k As Variant, i As Long, toRemove As Collection
    Set toRemove = New Collection
    k = d.Keys
    For i = LBound(k) To UBound(k)
        If Len(CStr(k(i))) < minLen Then toRemove.Add k(i)
    Next
    For i = 1 To toRemove.Count
        d.Remove toRemove(i)
    Next
End Sub
```

***

## 7) Performance Checklist (Excel)

*   **Read once, write once**: Convert ranges → arrays, process in memory, output arrays → range.
*   **Normalize** keys just once, consistently.
*   **Set `CompareMode`** before populating (switching later can silently fail).
*   **Avoid per‑cell worksheet calls** inside loops.
*   **Use `Value2`** to skip unnecessary type conversions.
*   **Prefer `RemoveAll`** over re‑instantiating in tight loops (reduces garbage).

***

## 8) Troubleshooting & Edge Cases

*   **Getting a missing key raises an error**: Always guard with `Exists`.
*   **Renaming keys (`dict.Key(old)=new`)** works but can be brittle (collisions). Safer: remove/add.
*   **Numbers and strings can collide** depending on how you normalize. Decide and be consistent (e.g., always `CStr` for keys).
*   **Dates as keys**: Either convert consistently (e.g., `CLng(CDate(...))`) or format as ISO strings (`Format(dt, "yyyy-mm-dd hh:nn:ss")`).
*   **Errors in cells**: detect with `IsError` and decide whether to skip or coerce.

***

## 9) Quick “Recipes” Index

*   **Mapping 2 columns**: `DictFromTwoColumns`
*   **Counts**: `FrequencyFromColumn`
*   **Grouping (one-to-many)**: `DictGroup`
*   **VLOOKUP alternative**: `FillPricesWithDict`
*   **Merge & resolve**: `MergeDicts`
*   **Sort keys/values**: `ReportSortedKeys` + `QuickSortKV`
*   **Header index**: `HeaderIndexMap`
*   **Composite keys**: `DictFromComposite`
*   **Unique list (ordered)**: `UniquePreserveOrder`

***

## 10) Optional: Early Binding Variants

If you set **Microsoft Scripting Runtime** reference, you can replace `As Object` with `As Scripting.Dictionary`:

```vb
Public Function NewDictEB(Optional ByVal caseInsensitive As Boolean = True) As Scripting.Dictionary
    Dim d As New Scripting.Dictionary
    d.CompareMode = IIf(caseInsensitive, vbTextCompare, vbBinaryCompare)
    Set NewDictEB = d
End Function
```

***

# Error Handling

Great ask. Below I’ve added a focused **Error Handling Toolkit** for `Scripting.Dictionary` work in VBA—especially when dictionaries are built from ranges. You’ll get:

*   Practical **patterns** (guard clauses, structured `On Error GoTo`, and “Try\*” functions that return `Boolean`).
*   **Robust wrappers** for common tasks (building from ranges, safe add/merge/rename, VLOOKUP-like fills).
*   **Validation helpers** for ranges and inputs.
*   A concise **reference table** of frequent dictionary-related errors.

> These drop into the same module as the earlier guide. They’re written to be copy‑ready.

***

## 1) Error Handling Patterns (Quick Summary)

| Pattern                              | When to Use                                            | Key Ideas                                                                              |
| ------------------------------------ | ------------------------------------------------------ | -------------------------------------------------------------------------------------- |
| Guard clauses                        | Validate inputs up front                               | Check `Nothing`/empty ranges, column counts, used area; exit early.                    |
| Structured handler (`On Error GoTo`) | Any proc that touches worksheet/range/dictionary       | Centralize cleanup, log `Err.Number/Source/Description`, decide to rethrow or swallow. |
| “Try\*” boolean functions            | When caller needs a safe attempt without runtime error | Return `True/False`, set `ByRef` results; no exceptions bubble up.                     |
| Resume Next + selective checks       | For expected, narrow failures                          | Use sparingly; immediately validate `Err.Number`; clear `Err`.                         |
| Custom error raising                 | When a precondition fails                              | `Err.Raise vbObjectError + N, "ProcName", "Message"` with actionable context.          |

***

## 2) Common Dictionary-Related Errors (Reference)

| Err.Number | When it happens                             | Typical Cause                                 | Fix                                          |
| ---------: | ------------------------------------------- | --------------------------------------------- | -------------------------------------------- |
|          5 | Invalid procedure call or argument          | Wrong type for key, invalid range sizing      | Normalize/validate, check range bounds       |
|          9 | Subscript out of range                      | Array bounds misuse                           | Always derive bounds via `LBound/UBound`     |
|         13 | Type mismatch                               | Unexpected cell Error values; non‑string keys | Coerce with `CStr`, check `IsError`          |
|         91 | Object variable not set                     | Using a `Nothing` dictionary                  | Ensure factory returns a valid object        |
|       1004 | Application-defined or object-defined error | Worksheet/range issues                        | Validate ranges; avoid per‑cell writing      |
|        457 | Key already associated with an element      | `dict.Add` with duplicate key                 | Use `Exists` or `dict(key) = value` (upsert) |

***

## 3) Guarded, Structured Error Handling Templates

### 3.1 Validate a Range (Guard Clause)

```vb
Public Function IsValid2ColRange(ByVal rng As Range, Optional ByRef why As String) As Boolean
    If rng Is Nothing Then
        why = "Range is Nothing.": Exit Function
    End If
    If rng.Rows.Count = 0 Or rng.Columns.Count < 2 Then
        why = "Range must have at least 2 columns and 1 row.": Exit Function
    End If
    If Application.WorksheetFunction.CountA(rng) = 0 Then
        why = "Range is empty.": Exit Function
    End If
    IsValid2ColRange = True
End Function
```

***

### 3.2 Centralized Error Reporter (Reusable)

```vb
Public Sub ReportError(ByVal procName As String)
    Debug.Print "ERROR in " & procName & _
                " | #" & Err.Number & " | " & Err.Source & " | " & Err.Description
End Sub
```

***

### 3.3 Structured Handler with Cleanup

```vb
Public Sub ExampleWithCleanup()
    Const PROC As String = "ExampleWithCleanup"
    On Error GoTo CleanFail
    Dim calc As XlCalculation
    calc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' ... your logic here ...
    
CleanExit:
    Application.Calculation = calc
    Application.ScreenUpdating = True
    Exit Sub
CleanFail:
    ReportError PROC
    Resume CleanExit
End Sub
```

***

## 4) “Try\*” Functions (Boolean + ByRef) — No Exceptions Bubble Up

### 4.1 Try to Build Dictionary from Two Columns

```vb
Public Function TryDictFromTwoColumns(ByVal rng As Range, _
                                      ByRef dictOut As Object, _
                                      Optional ByVal caseInsensitive As Boolean = True, _
                                      Optional ByRef why As String) As Boolean
    Const PROC As String = "TryDictFromTwoColumns"
    On Error GoTo Fail
    
    Dim ok As Boolean
    ok = IsValid2ColRange(rng, why)
    If Not ok Then GoTo FailNoReport
    
    Dim d As Object: Set d = NewDict(caseInsensitive)
    Dim arr As Variant, r As Long, key As String, val As Variant
    arr = RngToArray(rng)
    If IsEmpty(arr) Then
        why = "Range resolved to empty array.": GoTo FailNoReport
    End If
    
    For r = LBound(arr, 1) To UBound(arr, 1)
        key = NormalizeKey(arr(r, 1))
        val = arr(r, 2)
        If Len(key) > 0 Then
            ' Upsert style avoids Err 457 (duplicate key)
            d(key) = val
        End If
    Next
    
    Set dictOut = d
    TryDictFromTwoColumns = True
    Exit Function
    
Fail:
    ReportError PROC
FailNoReport:
    TryDictFromTwoColumns = False
End Function
```

**Usage**

```vb
Sub DemoTryDict()
    Dim d As Object, why As String
    If TryDictFromTwoColumns(Sheets("Data").Range("A2:B500"), d, True, why) Then
        Debug.Print "Built dictionary with", d.Count, "items."
    Else
        Debug.Print "Failed: "; why
    End If
End Sub
```

***

### 4.2 Try Fill (Dictionary-Powered VLOOKUP), Robust to Errors in Source Range

```vb
Public Function TryFillWithDict(ByVal mapRng As Range, _
                                ByVal codesRng As Range, _
                                ByVal outRng As Range, _
                                Optional ByRef why As String) As Boolean
    Const PROC As String = "TryFillWithDict"
    On Error GoTo Fail
    
    Dim d As Object
    If Not TryDictFromTwoColumns(mapRng, d, True, why) Then GoTo FailNoReport
    
    Dim arrIn As Variant, arrOut() As Variant
    arrIn = RngToArray(codesRng)
    If IsEmpty(arrIn) Then
        why = "codesRng is empty.": GoTo FailNoReport
    End If
    
    Dim n As Long, i As Long, key As String
    n = UBound(arrIn, 1): ReDim arrOut(1 To n, 1 To 1)
    
    For i = 1 To n
        If IsError(arrIn(i, 1)) Then
            arrOut(i, 1) = CVErr(xlErrValue)
        Else
            key = NormalizeKey(arrIn(i, 1))
            If Len(key) = 0 Then
                arrOut(i, 1) = CVErr(xlErrNA)
            ElseIf d.Exists(key) Then
                arrOut(i, 1) = d(key)
            Else
                arrOut(i, 1) = CVErr(xlErrNA)
            End If
        End If
    Next
    
    ArrayToRange outRng, arrOut
    TryFillWithDict = True
    Exit Function
    
Fail:
    ReportError PROC
FailNoReport:
    TryFillWithDict = False
End Function
```

***

## 5) Safe Wrappers for Specific Dictionary Operations

### 5.1 Safe Add (Avoid Err 457)

```vb
Public Function DictSafeAdd(ByVal d As Object, ByVal key As Variant, ByVal value As Variant) As Boolean
    On Error GoTo Fail
    Dim k As String: k = NormalizeKey(key)
    If Len(k) = 0 Then Exit Function
    If d.Exists(k) Then
        DictSafeAdd = False ' already present
    Else
        d.Add k, value
        DictSafeAdd = True
    End If
    Exit Function
Fail:
    DictSafeAdd = False
End Function
```

***

### 5.2 Safe Upsert

```vb
Public Sub DictUpsert(ByVal d As Object, ByVal key As Variant, ByVal value As Variant)
    On Error GoTo Fail
    Dim k As String: k = NormalizeKey(key)
    If Len(k) = 0 Then Exit Sub
    d(k) = value
    Exit Sub
Fail:
    ' Optional: log
    ReportError "DictUpsert"
End Sub
```

***

### 5.3 Safe Merge with Conflict Policy (Prefer Left/Right)

```vb
Public Function TryMergeDicts(ByVal leftD As Object, ByVal rightD As Object, _
                              ByRef merged As Object, _
                              Optional ByVal preferRight As Boolean = True, _
                              Optional ByRef why As String) As Boolean
    Const PROC As String = "TryMergeDicts"
    On Error GoTo Fail
    If leftD Is Nothing Or rightD Is Nothing Then
        why = "Either leftD or rightD is Nothing.": GoTo FailNoReport
    End If
    
    Dim d As Object: Set d = NewDict(True)
    Dim k As Variant, i As Long
    
    k = leftD.Keys
    For i = LBound(k) To UBound(k)
        d(k(i)) = leftD(k(i))
    Next
    
    k = rightD.Keys
    For i = LBound(k) To UBound(k)
        If d.Exists(k(i)) Then
            If preferRight Then d(k(i)) = rightD(k(i))
        Else
            d(k(i)) = rightD(k(i))
        End If
    Next
    
    Set merged = d
    TryMergeDicts = True
    Exit Function
    
Fail:
    ReportError PROC
FailNoReport:
    TryMergeDicts = False
End Function
```

***

### 5.4 Safe Rename Key (Prevent Collision)

```vb
Public Function TryRenameKey(ByVal d As Object, ByVal oldKey As Variant, ByVal newKey As Variant, _
                             Optional ByRef why As String) As Boolean
    Const PROC As String = "TryRenameKey"
    On Error GoTo Fail
    If d Is Nothing Then
        why = "Dictionary is Nothing.": GoTo FailNoReport
    End If
    Dim okOld As String: okOld = NormalizeKey(oldKey)
    Dim okNew As String: okNew = NormalizeKey(newKey)
    If Len(okOld) = 0 Or Len(okNew) = 0 Then
        why = "Keys cannot be empty.": GoTo FailNoReport
    End If
    If Not d.Exists(okOld) Then
        why = "Old key does not exist: " & okOld: GoTo FailNoReport
    End If
    If d.Exists(okNew) Then
        why = "New key already exists: " & okNew: GoTo FailNoReport
    End If
    
    ' Either use Key property or remove/add:
    d.Key(okOld) = okNew
    TryRenameKey = True
    Exit Function
    
Fail:
    ReportError PROC
FailNoReport:
    TryRenameKey = False
End Function
```

***

## 6) Robust Range → Dictionary Builders (With Error Handling)

### 6.1 Two-Column Builder with Explicit Error Path

```vb
Public Function DictFromTwoColumnsSafe(ByVal rng As Range, _
                                       Optional ByVal caseInsensitive As Boolean = True) As Object
    Const PROC As String = "DictFromTwoColumnsSafe"
    On Error GoTo Fail
    
    Dim why As String
    If Not IsValid2ColRange(rng, why) Then
        Err.Raise vbObjectError + 101, PROC, "Invalid input: " & why
    End If
    
    Dim d As Object: Set d = NewDict(caseInsensitive)
    Dim arr As Variant, r As Long, k As String, v As Variant
    arr = RngToArray(rng)
    If IsEmpty(arr) Then
        Err.Raise vbObjectError + 102, PROC, "Empty array from range."
    End If
    
    For r = LBound(arr, 1) To UBound(arr, 1)
        If Not IsError(arr(r, 1)) Then
            k = NormalizeKey(arr(r, 1))
            If Len(k) > 0 Then d(k) = arr(r, 2)
        End If
    Next
    
    Set DictFromTwoColumnsSafe = d
    Exit Function
    
Fail:
    ReportError PROC
    ' Re-raise so callers can trap if they want:
    Err.Raise Err.Number, PROC, Err.Description
End Function
```

***

### 6.2 Composite Key Builder with Defensive Checks

```vb
Public Function DictFromCompositeSafe(ByVal rng As Range) As Object
    Const PROC As String = "DictFromCompositeSafe"
    On Error GoTo Fail
    
    If rng Is Nothing Then Err.Raise vbObjectError + 201, PROC, "Range is Nothing."
    If rng.Columns.Count < 2 Then Err.Raise vbObjectError + 202, PROC, "Need >= 2 columns."
    
    Dim d As Object: Set d = NewDict(True)
    Dim arr As Variant: arr = RngToArray(rng)
    If IsEmpty(arr) Then Err.Raise vbObjectError + 203, PROC, "Range is empty."
    
    Dim r As Long, c As Long, lastCol As Long
    lastCol = UBound(arr, 2)
    
    For r = LBound(arr, 1) To UBound(arr, 1)
        Dim parts() As Variant
        ReDim parts(1 To lastCol - 1)
        For c = 1 To lastCol - 1
            If IsError(arr(r, c)) Then parts(c) = vbNullString Else parts(c) = arr(r, c)
        Next
        Dim key As String: key = MakeKey(parts)
        If Len(key) > 0 Then
            d(key) = arr(r, lastCol)
        End If
    Next
    
    Set DictFromCompositeSafe = d
    Exit Function
Fail:
    ReportError PROC
    Err.Raise Err.Number, PROC, Err.Description
End Function
```

***

## 7) Handling Cell Errors and Blanks Explicitly

### 7.1 Skip or Map Excel Errors

```vb
Public Sub DictFromTwoColumns_HandleCellErrors(ByVal src As Range, ByRef d As Object)
    Const PROC As String = "DictFromTwoColumns_HandleCellErrors"
    On Error GoTo Fail
    If d Is Nothing Then Set d = NewDict(True)
    
    Dim arr As Variant: arr = RngToArray(src)
    Dim r As Long, k As String, v As Variant
    
    For r = LBound(arr, 1) To UBound(arr, 1)
        If Not IsError(arr(r, 1)) Then
            k = NormalizeKey(arr(r, 1))
            If Len(k) > 0 Then
                If IsError(arr(r, 2)) Then
                    v = CVErr(xlErrNA) ' or vbNullString
                Else
                    v = arr(r, 2)
                End If
                d(k) = v
            End If
        End If
    Next
    Exit Sub
Fail:
    ReportError PROC
End Sub
```

***

## 8) Defensive Iteration and Removal

### 8.1 Remove by Predicate with Full Safety

```vb
Public Function TryRemoveByPredicate(ByVal d As Object, ByRef removedCount As Long, _
                                     ByVal minLen As Long, _
                                     Optional ByRef why As String) As Boolean
    Const PROC As String = "TryRemoveByPredicate"
    On Error GoTo Fail
    removedCount = 0
    If d Is Nothing Then
        why = "Dictionary is Nothing.": GoTo FailNoReport
    End If
    Dim k As Variant, i As Long
    k = d.Keys
    For i = LBound(k) To UBound(k)
        If Len(CStr(k(i))) < minLen Then
            d.Remove k(i)
            removedCount = removedCount + 1
        End If
    Next
    TryRemoveByPredicate = True
    Exit Function
Fail:
    ReportError PROC
FailNoReport:
    TryRemoveByPredicate = False
End Function
```

***

## 9) Logging & Escalation Strategy

*   **In library code**: Prefer `Try*` functions that return `Boolean` and set a `why` message, so callers can decide what to do.
*   **In app code (macros)**: Use structured handlers, log via `ReportError`, optionally **re-raise** (`Err.Raise`) if the error should abort the workflow.
*   When you **change application state** (screen updating, calc mode), always **restore in `CleanExit`** even on failure (see §3.3).

***

## 10) End-to-End Example: Safe De‑duplication Macro

This ties together validation, structured error handling, and clean up.

```vb
Public Sub UniqueToReport_Safe()
    Const PROC As String = "UniqueToReport_Safe"
    On Error GoTo CleanFail
    
    Dim wsIn As Worksheet, wsOut As Worksheet
    Set wsIn = ThisWorkbook.Worksheets("Data")
    Set wsOut = ThisWorkbook.Worksheets("Report")
    
    Dim src As Range: Set src = wsIn.Range("A2", wsIn.Cells(wsIn.Rows.Count, "A").End(xlUp))
    Dim why As String, ok As Boolean
    If src Is Nothing Or Application.WorksheetFunction.CountA(src) = 0 Then
        Err.Raise vbObjectError + 301, PROC, "Source range empty."
    End If
    
    Dim d As Object: Set d = NewDict(True)
    Dim arr As Variant: arr = RngToArray(src)
    Dim r As Long, k As String
    
    For r = LBound(arr, 1) To UBound(arr, 1)
        If Not IsError(arr(r, 1)) Then
            k = NormalizeKey(arr(r, 1))
            If Len(k) > 0 Then
                If Not d.Exists(k) Then d.Add k, arr(r, 1) ' preserve original text/case
            End If
        End If
    Next
    
    If d.Count = 0 Then
        Err.Raise vbObjectError + 302, PROC, "No unique values produced."
    End If
    
    ' Dump results
    Dim n As Long: n = d.Count
    Dim v As Variant: v = d.Items
    Dim out() As Variant, i As Long
    ReDim out(1 To n, 1 To 1)
    For i = 1 To n
        out(i, 1) = v(i - 1)
    Next
    ArrayToRange wsOut.Range("B2"), out
    
CleanExit:
    Exit Sub
CleanFail:
    ReportError PROC
    Resume CleanExit
End Sub
```


---

[DOC MOC](./doc-00_MOC.md)