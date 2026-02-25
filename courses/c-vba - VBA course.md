# VBA Course

- [VBA Course](#vba-course)
  - [Overview](#overview)
    - [Learning Objectives](#learning-objectives)
  - [Module 1 — Introduction to VBA Development (Environment, Inspector \& Debugging)](#module-1--introduction-to-vba-development-environment-inspector--debugging)
  - [Module 2 — VBA Syntax Basics (Core Building Blocks)](#module-2--vba-syntax-basics-core-building-blocks)
  - [Module 3 — Operating with the Spreadsheet (Bread‑and‑Butter Tasks)](#module-3--operating-with-the-spreadsheet-breadandbutter-tasks)
    - [3.1 Sheet Operations \& Inspection](#31-sheet-operations--inspection)
    - [3.2 Operations on Cells: Selection \& Write Values](#32-operations-on-cells-selection--write-values)
  - [Module 4 — Patterns for Engineering Data Automation](#module-4--patterns-for-engineering-data-automation)
  - [Module 5 — Engineering Examples (Data Automation \& Physical Simulation)](#module-5--engineering-examples-data-automation--physical-simulation)
    - [Example 1 — Data Automation: Consolidate CSV‑like Sheets to a Master Table](#example-1--data-automation-consolidate-csvlike-sheets-to-a-master-table)
    - [Example 2 — Physical Simulation: Pipe Pressure Drop (Darcy–Weisbach with Swamee–Jain)](#example-2--physical-simulation-pipe-pressure-drop-darcyweisbach-with-swameejain)
    - [Example 3 — Engineering Design: Control Valve Cv Sizing (Liquids)](#example-3--engineering-design-control-valve-cv-sizing-liquids)
  - [Module 6 — Robustness, Testing, and Traceability](#module-6--robustness-testing-and-traceability)
  - [Module 7 — Performance \& Interop Tips](#module-7--performance--interop-tips)
  - [Module 8 — Capstone Exercise](#module-8--capstone-exercise)
  - [Suggested Schedule](#suggested-schedule)
  - [What you’ll take away](#what-youll-take-away)
    - [Quick Next Steps](#quick-next-steps)
- [Engineering Examples](#engineering-examples)
  - [1) HVAC Hydraulics](#1-hvac-hydraulics)
    - [Summary (use case quick scan)](#summary-use-case-quick-scan)
    - [1.1 Duct pressure drop (Darcy–Weisbach + Haaland)](#11-duct-pressure-drop-darcyweisbach--haaland)
    - [1.2 Hydronic loop balancing (simple valve position heuristic)](#12-hydronic-loop-balancing-simple-valve-position-heuristic)
    - [1.3 Pump sizing \& motor selection](#13-pump-sizing--motor-selection)
  - [2) Test Data Consolidation](#2-test-data-consolidation)
    - [Summary](#summary)
    - [2.1 Import all CSVs in folder into a MASTER table](#21-import-all-csvs-in-folder-into-a-master-table)
    - [2.2 Build KPIs (PivotTable: Mean/StDev by Field and Source)](#22-build-kpis-pivottable-meanstdev-by-field-and-source)
    - [2.3 Resample a time series to uniform Δt (linear interpolation)](#23-resample-a-time-series-to-uniform-δt-linear-interpolation)
  - [3) Thermodynamics](#3-thermodynamics)
    - [Summary](#summary-1)
    - [3.1 Antoine equation utilities (Psat and Boiling Point)](#31-antoine-equation-utilities-psat-and-boiling-point)
    - [3.2 Heat exchanger (counterflow) using NTU–ε](#32-heat-exchanger-counterflow-using-ntuε)
    - [3.3 Simplified psychrometrics (Tdb, RH, P → W, h)](#33-simplified-psychrometrics-tdb-rh-p--w-h)
  - [4) Materials](#4-materials)
    - [Summary](#summary-2)
    - [4.1 Von Mises stress (general 3D)](#41-von-mises-stress-general-3d)
    - [4.2 Beam deflection \& bending stress (simply supported, center point load)](#42-beam-deflection--bending-stress-simply-supported-center-point-load)
    - [4.3 Materials screening by Figure‑of‑Merit (FoM)](#43-materials-screening-by-figureofmerit-fom)
  - [How to deploy these quickly](#how-to-deploy-these-quickly)
  - [Would you like these wrapped into a starter workbook?](#would-you-like-these-wrapped-into-a-starter-workbook)
- [Error Handling](#error-handling)
  - [0) Reusable Error Utilities (drop into `modErrorUtils`)](#0-reusable-error-utilities-drop-into-moderrorutils)
  - [1) Patterns You’ll Reuse Everywhere](#1-patterns-youll-reuse-everywhere)
    - [1.1 Structured handler with “finally”](#11-structured-handler-with-finally)
    - [1.2 Guarded “probe” with `Resume Next` (safe only for *expected* failures)](#12-guarded-probe-with-resume-next-safe-only-for-expected-failures)
    - [1.3 Retry pattern (e.g., transient file locks)](#13-retry-pattern-eg-transient-file-locks)
    - [1.4 Developer‑only checks](#14-developeronly-checks)
  - [2) HVAC Hydraulics — Error‑Hardened Examples](#2-hvac-hydraulics--errorhardened-examples)
    - [2.1 Duct Pressure Drop (with validation \& context)](#21-duct-pressure-drop-with-validation--context)
    - [2.2 Hydronic Balancing (guard numeric issues)](#22-hydronic-balancing-guard-numeric-issues)
    - [2.3 Pump Power \& Motor Selection (with bounds and messages)](#23-pump-power--motor-selection-with-bounds-and-messages)
  - [3) Test Data Consolidation — Error‑Hardened Examples](#3-test-data-consolidation--errorhardened-examples)
    - [3.1 Import CSV Folder → MASTER (with retry + summary)](#31-import-csv-folder--master-with-retry--summary)
    - [3.2 Build KPIs (Pivot) — handle caching errors](#32-build-kpis-pivot--handle-caching-errors)
  - [4) Thermodynamics — Error‑Hardened Examples](#4-thermodynamics--errorhardened-examples)
    - [4.1 Antoine Utilities (parameter sanity + bracket failure)](#41-antoine-utilities-parameter-sanity--bracket-failure)
    - [4.2 Heat Exchanger NTU‑ε (guard Cr≈1 and invalid inputs)](#42-heat-exchanger-ntuε-guard-cr1-and-invalid-inputs)
  - [5) Materials — Error‑Hardened Examples](#5-materials--errorhardened-examples)
    - [5.1 Von Mises (ensure numeric \& pass/fail labeling)](#51-von-mises-ensure-numeric--passfail-labeling)
    - [5.2 Beam (validate geometry \& modulus)](#52-beam-validate-geometry--modulus)
  - [6) Additional Patterns You Might Need](#6-additional-patterns-you-might-need)
    - [6.1 Safe `.Find` usage (avoids hidden state issues)](#61-safe-find-usage-avoids-hidden-state-issues)
    - [6.2 Safe array write/read with dimension checks](#62-safe-array-writeread-with-dimension-checks)
    - [6.3 User‑facing consolidated messages after batch](#63-userfacing-consolidated-messages-after-batch)
  - [7) How to integrate quickly](#7-how-to-integrate-quickly)

---

## Overview

Below is a compact, engineering‑oriented Excel VBA course outline with practical labs and ready‑to‑run snippets. It’s structured to take participants from zero to productive, with an emphasis on robust patterns for automation and engineering calculations.

***

**Audience:** Engineers and technical staff who need to automate data handling, calculations, and reporting in Excel.  
**Format:** 1–2 days (modular), hands‑on labs, templates, and re‑usable code patterns.  
**Prerequisites:** Proficiency with Excel; basic programming concepts helpful but not required.

### Learning Objectives

*   Set up a productive VBA development environment and use debugging/inspection tools effectively.
*   Master essential VBA syntax (loops, conditionals, procedures, error handling).
*   Work fluently with the Excel Object Model (Workbooks/Worksheets/Range).
*   Build robust, fast automations for engineering data workflows.
*   Implement engineering calculations (e.g., pressure drop, valve sizing) with readable, testable VBA.

***

## Module 1 — Introduction to VBA Development (Environment, Inspector & Debugging)

**Topics**

*   **Enable Developer tab:** File → Options → Customize Ribbon → Developer.
*   **Macro security:** File → Options → Trust Center → Trust Center Settings → Macro Settings (set according to org policy).
*   **VBE tour:**
    *   **Project Explorer** (Ctrl+R) – structure of workbooks/modules/forms.
    *   **Properties Window** (F4) – object properties.
    *   **Code Window** – editor, procedures list.
    *   **Immediate Window** (Ctrl+G) – quick evaluation `?ActiveSheet.Name`, `Debug.Print`, one‑off commands.
    *   **Locals Window** – inspect variables in current scope.
    *   **Watch Window** – track expressions; break when value changes.
    *   **Object Browser** (F2) – discover classes, methods, constants.
*   **Debugging workflow:**
    *   Set breakpoints (F9), Step Into (F8), Step Over (Shift+F8), Step Out (Ctrl+Shift+F8).
    *   Use `Debug.Print`, `Stop`, `Debug.Assert <boolean>`.
    *   Error handling pattern with `On Error GoTo`.

**Lab 1**

*   Open VBE, write a “Hello, workbook” macro, set a breakpoint, step through, inspect variables in Locals/Watch, print into Immediate.

***

## Module 2 — VBA Syntax Basics (Core Building Blocks)

> The table gives quick patterns. Detailed examples follow in code blocks.

| Purpose              | Statement                                | Inline Syntax Pattern                                                          | Notes                                        |
| -------------------- | ---------------------------------------- | ------------------------------------------------------------------------------ | -------------------------------------------- |
| Variable declaration | `Dim`, `Const`                           | `Dim x As Double`, `Const PI As Double = 3.14159`                              | Prefer explicit types and `Option Explicit`. |
| Procedures           | `Sub`, `Function`                        | `Sub Name()` … `End Sub`; `Function F(a As Double) As Double` … `End Function` | Functions return values; Subs don’t.         |
| Conditional          | `If…Then…ElseIf…Else…End If`             | `If a > 0 Then … Else … End If`                                                | Nested logic; keep branches short/readable.  |
| Multi‑branch         | `Select Case`                            | `Select Case x` `Case 1` … `Case Else` … `End Select`                          | Clean alternative to many `ElseIf`.          |
| Loop (counter)       | `For…Next`                               | `For i = 1 To n` … `Next i`                                                    | Use `Exit For` to break.                     |
| Loop (collection)    | `For Each…Next`                          | `For Each ws In Worksheets` … `Next`                                           | Ideal for ranges, sheets, arrays.            |
| Loop (conditional)   | `Do While…Loop` / `Do Until…Loop`        | `Do While cond` … `Loop`                                                       | Guard against infinite loops.                |
| With block           | `With…End With`                          | `With Range("A1")` `.Value = 1` `End With`                                     | Reduces repeated qualifiers.                 |
| Error handling       | `On Error GoTo` / `On Error Resume Next` | `On Error GoTo ErrH`                                                           | Prefer structured handler + cleanup.         |
| Exit                 | `Exit For/Do/Sub/Function`               | `If done Then Exit For`                                                        | Use judiciously to keep code clear.          |
| Arrays               | `Dim a()` / `ReDim`                      | `ReDim a(1 To n)`                                                              | For batch read/write with ranges.            |
| String ops           | `&`, `Replace`, `Split`, `Join`          | `s = "a" & "b"`                                                                | Useful for parsing text data.                |
| Date/Time            | `Now`, `DateAdd`, `Timer`                | `t0 = Timer`                                                                   | For timestamps and simple timing.            |

**Code examples**

```vb
Option Explicit

' Function example
Function FahrenheitToCelsius(F As Double) As Double
    FahrenheitToCelsius = (F - 32#) * 5# / 9#
End Function

' Loop and conditional example
Sub SumPositiveValues()
    Dim total As Double, c As Range
    For Each c In Range("B2:B100")
        If IsNumeric(c.Value2) And c.Value2 > 0 Then
            total = total + c.Value2
        End If
    Next
    Range("B101").Value = total
End Sub

' Error handling template
Sub RobustTemplate()
    On Error GoTo ErrH
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' ... your logic ...

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
ErrH:
    Debug.Print "Error " & Err.Number & ": " & Err.Description
    Resume CleanExit
End Sub
```

***

## Module 3 — Operating with the Spreadsheet (Bread‑and‑Butter Tasks)

### 3.1 Sheet Operations & Inspection

**Retrieve all sheet names (write to an index sheet)**

```vb
Sub ListSheetNames()
    Dim ws As Worksheet, i As Long
    Dim idx As Worksheet
    
    On Error Resume Next
    Set idx = ThisWorkbook.Worksheets("Index")
    On Error GoTo 0
    If idx Is Nothing Then
        Set idx = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        idx.Name = "Index"
    Else
        idx.Cells.ClearContents
    End If

    idx.Range("A1").Value = "Sheet Name"
    i = 2
    For Each ws In ThisWorkbook.Worksheets
        idx.Cells(i, 1).Value = ws.Name
        i = i + 1
    Next ws
End Sub
```

**Create / Copy / Move / Activate a sheet (best practice: qualify objects)**

```vb
Sub CreateCopyMoveActivate()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim src As Worksheet, dst As Worksheet
    
    ' Create a new sheet at the end
    Set dst = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    dst.Name = "Scratch"

    ' Copy an existing sheet
    Set src = wb.Worksheets(1)
    src.Copy After:=dst  ' creates a new sheet copy

    ' Move the copied sheet to the first position
    wb.Worksheets(wb.Worksheets.Count).Move Before:=wb.Worksheets(1)

    ' Activate a sheet (rarely needed; see note below)
    wb.Worksheets("Scratch").Activate
End Sub
```

> **Note:** Prefer working with objects directly rather than `Activate/Select`. It’s faster and more reliable in automation.

**Check if a sheet exists**

```vb
Function SheetExists(sheetName As String, Optional wb As Workbook) As Boolean
    Dim ws As Worksheet
    If wb Is Nothing Then Set wb = ThisWorkbook
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then
            SheetExists = True
            Exit Function
        End If
    Next
End Function
```

### 3.2 Operations on Cells: Selection & Write Values

**Avoid `Select` and use fully‑qualified references**

```vb
Sub WriteWithoutSelect()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet: Set ws = wb.Worksheets("Data")
    Dim rng As Range: Set rng = ws.Range("A2").Resize(5, 2)
    
    ' Write a 2D array to a range in one go (fast)
    Dim arr(1 To 5, 1 To 2) As Variant
    Dim i As Long
    For i = 1 To 5
        arr(i, 1) = "Item-" & i
        arr(i, 2) = i * 10
    Next
    rng.Value = arr  ' batch write
    
    ' Read back to array (fast)
    Dim back As Variant
    back = rng.Value
    Debug.Print "First item read-back:", back(1, 1)
End Sub
```

**Common patterns**

```vb
Sub CellPatterns()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(1)

    ' Single cell and formula
    ws.Range("D2").Value = "Mass [kg]"
    ws.Range("E2").Formula = "=D2*9.81"  ' Newtons

    ' Offset and Resize
    ws.Range("B2").Offset(0, 2).Value = "Shifted"
    ws.Range("A10").Resize(3, 3).ClearContents

    ' Cells(row, col)
    ws.Cells(5, 1).Value = "Row5Col1"

    ' Find
    Dim f As Range
    Set f = ws.Cells.Find(What:="Target", LookIn:=xlValues, LookAt:=xlWhole)
    If Not f Is Nothing Then f.Interior.Color = vbYellow

    ' UsedRange
    Dim ur As Range
    Set ur = ws.UsedRange
    Debug.Print "Used rows:", ur.Rows.Count
End Sub
```

**Performance toggles (use sparingly and always restore)**

```vb
Sub SpeedBlockStart()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
End Sub

Sub SpeedBlockEnd()
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
```

**Lab 2**

*   Build a macro that reads a 2D block, filters by a threshold, writes results to a new sheet, and creates a header row with units.

***

## Module 4 — Patterns for Engineering Data Automation

*   **Pattern A: Batch read/write with arrays** (thousands of cells, minimal object calls).
*   **Pattern B: Table‑driven macros** (headers define where/what to compute, macro iterates rows).
*   **Pattern C: Logging & traceability** (`Debug.Print` and a “Log” sheet).
*   **Pattern D: Config‑first** (named ranges or a “Config” sheet for paths, constants, units).
*   **Pattern E: Robust error handling** (single exit, cleanup, helpful error messages).

**Lab 3**

*   Create a “Config” sheet for input directory, thresholds, and output location; macro reads config, processes data, and writes a summary table.

***

## Module 5 — Engineering Examples (Data Automation & Physical Simulation)

### Example 1 — Data Automation: Consolidate CSV‑like Sheets to a Master Table

*Scenario:* Multiple similarly structured sheets (or imported CSVs) with test data. Merge into one normalized table with a source tag and timestamp.

```vb
Option Explicit

Sub ConsolidateSheetsToMaster()
    On Error GoTo ErrH
    Application.ScreenUpdating = False
    
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet, master As Worksheet
    Dim nextRow As Long, data As Variant, lastRow As Long, lastCol As Long
    Dim src As String, tstamp As Date
    
    ' Prepare master sheet
    If Not SheetExists("MASTER", wb) Then wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count)).Name = "MASTER"
    Set master = wb.Worksheets("MASTER")
    master.Cells.ClearContents
    master.Range("A1:D1").Value = Array("SourceSheet", "Timestamp", "Parameter", "Value")
    nextRow = 2
    
    For Each ws In wb.Worksheets
        If ws.Name <> "MASTER" And ws.UsedRange.Rows.Count > 1 Then
            src = ws.Name
            tstamp = Now

            ' assume parameters in Col A and values in Col B (adjust as needed)
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            If lastRow >= 2 Then
                data = ws.Range("A2:B" & lastRow).Value ' 2D array
                Dim i As Long
                For i = 1 To UBound(data, 1)
                    master.Cells(nextRow, 1).Value = src
                    master.Cells(nextRow, 2).Value = tstamp
                    master.Cells(nextRow, 3).Value = data(i, 1)
                    master.Cells(nextRow, 4).Value = data(i, 2)
                    nextRow = nextRow + 1
                Next
            End If
        End If
    Next ws

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub
ErrH:
    Debug.Print "Consolidation error: "; Err.Number; Err.Description
    Resume CleanExit
End Sub
```

*Extensions:* Add unit normalization, create a PivotTable on “MASTER”, or export to CSV.

***

### Example 2 — Physical Simulation: Pipe Pressure Drop (Darcy–Weisbach with Swamee–Jain)

*Scenario:* Compute head loss or pressure drop for internal flow based on length, diameter, flow, roughness, and fluid properties.

```vb
Option Explicit

' Returns friction factor f using Swamee–Jain (valid for turbulent flow)
Private Function FricSwameeJain(Re As Double, relRough As Double) As Double
    If Re <= 0# Then
        FricSwameeJain = 0#
    Else
        FricSwameeJain = 0.25 / (Log10((relRough / 3.7) + (5.74 / (Re ^ 0.9)))) ^ 2
    End If
End Function

' Darcy–Weisbach pressure drop (Pa)
Public Function PressureDrop_DW( _
    ByVal rho As Double, _          ' density [kg/m^3]
    ByVal mu As Double, _           ' dynamic viscosity [Pa·s]
    ByVal Q As Double, _            ' volumetric flow [m^3/s]
    ByVal D As Double, _            ' diameter [m]
    ByVal L As Double, _            ' length [m]
    ByVal eps As Double) As Double  ' absolute roughness [m]

    Dim A As Double, V As Double, Re As Double, relR As Double, f As Double
    A = WorksheetFunction.Pi() * (D ^ 2) / 4#
    V = Q / A
    If D <= 0# Or A <= 0# Then
        PressureDrop_DW = 0#
        Exit Function
    End If
    Re = rho * V * D / mu
    relR = eps / D

    ' Laminar: f = 64/Re; Turbulent: Swamee–Jain
    If Re < 2300# Then
        f = 64# / Re
    Else
        f = FricSwameeJain(Re, relR)
    End If

    ' ΔP = f * (L/D) * (ρ V^2 / 2)
    PressureDrop_DW = f * (L / D) * (rho * V * V / 2#)
End Function

' Sheet macro: fill ΔP for each row based on inputs
Sub ComputePressureDropTable()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Hydraulics")
    Dim r As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For r = 2 To lastRow
        ws.Cells(r, "H").Value = PressureDrop_DW( _
            rho:=ws.Cells(r, "A").Value2, _
            mu:=ws.Cells(r, "B").Value2, _
            Q:=ws.Cells(r, "C").Value2, _
            D:=ws.Cells(r, "D").Value2, _
            L:=ws.Cells(r, "E").Value2, _
            eps:=ws.Cells(r, "F").Value2)
    Next
    ws.Range("H1").Value = "ΔP [Pa]"
End Sub
```

*Sheet layout suggestion (“Hydraulics”):*  
A: ρ \[kg/m³], B: μ \[Pa·s], C: Q \[m³/s], D: D \[m], E: L \[m], F: ε \[m], G: Notes, H: ΔP \[Pa].

***

### Example 3 — Engineering Design: Control Valve Cv Sizing (Liquids)

*Scenario:* Given a required flow, pressure drop, and specific gravity, compute required Cv for selection; map to nearest standard size.

```vb
Option Explicit

' Cv for liquids: Q = Cv * sqrt(ΔP / SG)  => Cv = Q / sqrt(ΔP / SG)
Public Function Cv_Required(ByVal Q_m3h As Double, ByVal dP_bar As Double, ByVal SG As Double) As Double
    Dim Q_gpm As Double, dP_psi As Double
    ' Convert to US customary for standard Cv relation if needed:
    ' 1 m^3/h = 4.402867 gpm ; 1 bar = 14.5038 psi
    Q_gpm = Q_m3h * 4.402867
    dP_psi = dP_bar * 14.5038
    If dP_psi <= 0# Or SG <= 0# Then
        Cv_Required = 0#
    Else
        Cv_Required = Q_gpm / Sqr(dP_psi / SG)
    End If
End Function

' Nearest standard Cv from a simple lookup table
Public Function Cv_SelectNearest(ByVal CvReq As Double) As Double
    Dim stdCv As Variant
    stdCv = Array(0.5, 1, 2, 3, 5, 8, 10, 15, 20, 30, 40, 60, 85, 110, 150, 200)
    Dim i As Long
    For i = LBound(stdCv) To UBound(stdCv)
        If CvReq <= stdCv(i) Then
            Cv_SelectNearest = stdCv(i)
            Exit Function
        End If
    Next
    Cv_SelectNearest = stdCv(UBound(stdCv)) ' largest
End Function

Sub ValveSizingTable()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("ValveSizing")
    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ws.Range("E1").Value = "Cv Required"
    ws.Range("F1").Value = "Cv Selected"
    
    For r = 2 To lastRow
        Dim CvReq As Double
        CvReq = Cv_Required( _
            Q_m3h:=ws.Cells(r, "A").Value2, _
            dP_bar:=ws.Cells(r, "B").Value2, _
            SG:=ws.Cells(r, "C").Value2)
        ws.Cells(r, "E").Value = CvReq
        ws.Cells(r, "F").Value = Cv_SelectNearest(CvReq)
    Next
End Sub
```

*Sheet layout suggestion (“ValveSizing”):*  
A: Flow \[m³/h], B: ΔP \[bar], C: SG \[–], D: Notes, E: Cv Required, F: Cv Selected.

***

## Module 6 — Robustness, Testing, and Traceability

*   **Structure:** One module per topic, clear naming (`modHydraulics`, `modValveSizing`).
*   **Input validation:** `IsNumeric`, bounds checks, units noted in headers.
*   **Logging:** Dedicated “Log” sheet; append timestamp, operation, count, warnings.
*   **Unit tests (lightweight):** Create a “Tests” module with assertion helpers:

```vb
Private Sub AssertAlmostEqual(actual As Double, expected As Double, Optional tol As Double = 1E-6, Optional msg As String = "")
    If Abs(actual - expected) > tol Then
        Err.Raise vbObjectError + 1000, "AssertAlmostEqual", "Mismatch: " & msg & " (act=" & actual & ", exp=" & expected & ")"
    End If
End Sub

Public Sub Test_PressureDrop()
    Dim dp As Double
    dp = PressureDrop_DW(1000#, 1E-3, 0.002, 0.05, 10, 1E-4)
    AssertAlmostEqual dp,  '<< set an expected regression value here
End Sub
```

***

## Module 7 — Performance & Interop Tips

*   **Minimize COM calls:** read/write ranges in blocks (arrays), not cell‑by‑cell.
*   **Avoid `Select/Activate`:** work on objects directly.
*   **Use `.Value2`** for speed/consistency.
*   **Turn off screen updating & events** around bulk ops; always restore in `CleanExit`.
*   **Named ranges & Tables (ListObjects):** stable references even as data grows.
*   **Interop:** When applicable, call Solver, Power Query refresh, or external tools (subject to policy).

***

## Module 8 — Capstone Exercise

*   **Brief:** Build a robust automation that:
    1.  consolidates test data,
    2.  runs an engineering calc (e.g., ΔP),
    3.  produces a summary table and a chart,
    4.  logs actions and timing,
    5.  exposes a user button on a clean worksheet UI.

Deliverables: Workbook, code module(s), and a README sheet.

***

## Suggested Schedule

*   **Half‑day (crash):** Modules 1–3 + Example 1
*   **One day (standard):** Modules 1–5 + Labs
*   **Two days (deep‑dive):** All modules + Capstone + code reviews

***

## What you’ll take away

*   Re‑usable templates (robust macro skeleton, error handling, speed blocks).
*   Engineering functions (pressure drop, valve Cv) ready to adapt to your domain.
*   Patterns for safe, maintainable automation aligned with engineering QA needs.

***

### Quick Next Steps

*   Which Excel version (and 32/64‑bit) will participants use, and are macros allowed in your environment?
*   Would you like the outline packaged as a slide deck or a Word/PDF handout with the labs?
*   Any specific engineering domains to emphasize (e.g., HVAC hydraulics, test data consolidation, thermodynamics, materials)?

# Engineering Examples

Great—here are **ready‑to‑run, engineering‑oriented VBA examples** across the four fields you asked for. Each field includes a short summary table (for quick scanning) and **clean, robust code** using best practices (`Option Explicit`, avoiding `Select/Activate`, error‑safe toggles). You can paste each module into a standard Excel `.xlsm` workbook.

***

## 1) HVAC Hydraulics

### Summary (use case quick scan)

| Example                             | Purpose                                                                  | Inputs (Sheet columns)                                                           | Outputs                                                |
| ----------------------------------- | ------------------------------------------------------------------------ | -------------------------------------------------------------------------------- | ------------------------------------------------------ |
| Duct pressure drop (segment list)   | Compute ΔP per duct segment (Darcy–Weisbach, Haaland) and total per path | `A:PathID`, `B:L[m]`, `C:D[m]`, `D:ε[m]`, `E:Q[m³/s]`, `F:ρ[kg/m³]`, `G:μ[Pa·s]` | `H:Re`, `I:f`, `J:V[m/s]`, `K:ΔP[Pa]`, per‑path totals |
| Hydronic loop balancing (heuristic) | Recommend new balancing valve position to meet design flow               | `A:Branch`, `B:Q_design`, `C:Q_measured`, `D:Pos_current[%]`                     | `E:Pos_new[%]`, `F:Note`                               |
| Pump sizing & motor selection       | Compute pump shaft power and select nearest standard motor               | `A:Q[m³/s]`, `B:H[m]`, `C:ρ[kg/m³]`, `D:η_pump[%]`                               | `E:P_shaft[kW]`, `F:Motor_kW`                          |

> **Sheets expected:** `"Ducts"`, `"Hydronic"`, `"Pumps"` (you can rename in code).

### 1.1 Duct pressure drop (Darcy–Weisbach + Haaland)

```vb
Option Explicit

Private Function FricHaaland(ByVal Re As Double, ByVal relR As Double) As Double
    ' Haaland explicit correlation: 1/sqrt(f) = -1.8 log10( ( (ε/D)/3.7 )^1.11 + 6.9/Re )
    If Re <= 0# Then
        FricHaaland = 0#
        Exit Function
    End If
    Dim invSqrtF As Double
    invSqrtF = -1.8 * Log10((relR / 3.7) ^ 1.11 + 6.9 / Re)
    FricHaaland = 1# / (invSqrtF ^ 2)
End Function

Public Sub DuctPressureDropSegments()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Ducts")
    Dim r As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ws.Range("H1:K1").Value = Array("Re", "f", "V [m/s]", "ΔP [Pa]")
    
    For r = 2 To lastRow
        Dim L As Double, D As Double, eps As Double, Q As Double, rho As Double, mu As Double
        Dim A As Double, V As Double, Re As Double, relR As Double, f As Double, dP As Double
        
        L = ws.Cells(r, "B").Value2
        D = ws.Cells(r, "C").Value2
        eps = ws.Cells(r, "D").Value2
        Q = ws.Cells(r, "E").Value2
        rho = ws.Cells(r, "F").Value2
        mu = ws.Cells(r, "G").Value2
        
        If D > 0# Then
            A = WorksheetFunction.Pi() * D ^ 2 / 4#
        Else
            A = 0#
        End If
        
        If A > 0# Then
            V = Q / A
        Else
            V = 0#
        End If
        
        If mu > 0# And D > 0# Then
            Re = rho * V * D / mu
        Else
            Re = 0#
        End If
        relR = IIf(D > 0#, eps / D, 0#)
        
        If Re < 2300# And Re > 0# Then
            f = 64# / Re
        Else
            f = FricHaaland(Re, relR)
        End If
        
        dP = f * (L / D) * (rho * V * V / 2#)
        
        ws.Cells(r, "H").Value = Re
        ws.Cells(r, "I").Value = f
        ws.Cells(r, "J").Value = V
        ws.Cells(r, "K").Value = dP
    Next
    
    ' Optional: total ΔP per PathID (col A) using a simple subtotal
    ws.Range("M1").Value = "PathID"
    ws.Range("N1").Value = "Total ΔP [Pa]"
    ws.Range("M2").Resize(lastRow - 1, 1).Value = ws.Range("A2:A" & lastRow).Value
    ws.Range("N2").Resize(lastRow - 1, 1).FormulaR1C1 = "=SUMIF(C1,RC[-1],C11)"
End Sub
```

### 1.2 Hydronic loop balancing (simple valve position heuristic)

```vb
Option Explicit

' Heuristic: Position_new = Position_current * (Q_design / Q_measured)^2 , clamped to [5, 100] %
Public Sub HydronicBalanceRecommend()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Hydronic")
    Dim r As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ws.Range("E1:F1").Value = Array("Pos_new [%]", "Note")
    
    For r = 2 To lastRow
        Dim Qd As Double, Qm As Double, posNow As Double, posNew As Double
        Qd = ws.Cells(r, "B").Value2
        Qm = ws.Cells(r, "C").Value2
        posNow = ws.Cells(r, "D").Value2
        
        If Qm > 0# And posNow > 0# Then
            posNew = posNow * (Qd / Qm) ^ 2
            If posNew < 5# Then posNew = 5#
            If posNew > 100# Then posNew = 100#
            ws.Cells(r, "E").Value = posNew
            ws.Cells(r, "F").Value = IIf(Abs(Qd - Qm) / Application.Max(Qd, 1E-9) < 0.05, "OK (<5% err)", "Adjust")
        Else
            ws.Cells(r, "E").Value = ""
            ws.Cells(r, "F").Value = "Missing/invalid inputs"
        End If
    Next
End Sub
```

### 1.3 Pump sizing & motor selection

```vb
Option Explicit

' P_shaft [kW] = ρ g Q H / η / 1000
Public Sub PumpPowerSelect()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Pumps")
    Dim r As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ws.Range("E1:F1").Value = Array("P_shaft [kW]", "Motor_kW (nearest)")
    
    Dim g As Double: g = 9.80665
    Dim motors As Variant
    motors = Array(0.37, 0.55, 0.75, 1.1, 1.5, 2.2, 3, 4, 5.5, 7.5, 11, 15, 18.5, 22, 30, 37, 45, 55, 75, 90, 110, 132)
    
    Dim i As Long
    For r = 2 To lastRow
        Dim Q As Double, H As Double, rho As Double, eta As Double, PkW As Double, sel As Double
        Q = ws.Cells(r, "A").Value2
        H = ws.Cells(r, "B").Value2
        rho = ws.Cells(r, "C").Value2
        eta = ws.Cells(r, "D").Value2 / 100#
        
        If Q > 0# And H > 0# And rho > 0# And eta > 0# Then
            PkW = rho * g * Q * H / eta / 1000#
            ws.Cells(r, "E").Value = PkW
            sel = motors(UBound(motors))
            For i = LBound(motors) To UBound(motors)
                If PkW <= motors(i) * 0.9 Then ' 10% margin
                    sel = motors(i)
                    Exit For
                End If
            Next
            ws.Cells(r, "F").Value = sel
        Else
            ws.Cells(r, "E").Value = ""
            ws.Cells(r, "F").Value = ""
        End If
    Next
End Sub
```

***

## 2) Test Data Consolidation

### Summary

| Example                          | Purpose                                                             | Inputs                                               | Outputs                                                         |
| -------------------------------- | ------------------------------------------------------------------- | ---------------------------------------------------- | --------------------------------------------------------------- |
| Import CSV folder → Master table | Batch import all CSVs in a folder and normalize into a master table | `"Config!B1": FolderPath`                            | `MASTER` sheet with `SourceFile`, `Timestamp`, columns from CSV |
| Create PivotTable KPIs           | Build basic KPIs (mean, stdev) across fields                        | `MASTER`                                             | New sheet `KPIs` with PivotTable                                |
| Resample time series             | Resample irregular time series to Δt (s) with linear interpolation  | Any sheet: `A:Time(s)`, `B:Value`; `"Config!B2": Δt` | New sheet with uniform grid                                     |

### 2.1 Import all CSVs in folder into a MASTER table

```vb
Option Explicit

Public Sub ImportCSVFolderToMaster()
    Dim cfg As Worksheet: Set cfg = ThisWorkbook.Worksheets("Config")
    Dim folder As String: folder = cfg.Range("B1").Value  ' e.g., C:\Data\Runs\
    If Len(folder) = 0 Then
        MsgBox "Set Config!B1 to the folder path.", vbExclamation
        Exit Sub
    End If
    
    Dim master As Worksheet
    On Error Resume Next
    Set master = ThisWorkbook.Worksheets("MASTER")
    On Error GoTo 0
    If master Is Nothing Then
        Set master = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        master.Name = "MASTER"
    Else
        master.Cells.Clear
    End If
    
    Dim nextRow As Long: nextRow = 2
    master.Range("A1:D1").Value = Array("SourceFile", "Timestamp", "Field", "Value")
    
    Dim f As String: f = Dir(folder & "*.csv")
    Application.ScreenUpdating = False
    Do While Len(f) > 0
        Dim wb As Workbook
        Dim srcPath As String: srcPath = folder & f
        Set wb = Workbooks.Open(Filename:=srcPath, ReadOnly:=True)
        
        Dim ws As Worksheet: Set ws = wb.Worksheets(1) ' assumes single sheet
        Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        Dim rng As Range: Set rng = ws.Range("A1:B" & lastRow) ' assumes A=Field, B=Value
        Dim arr As Variant: arr = rng.Value
        
        Dim i As Long
        For i = 2 To UBound(arr, 1)
            master.Cells(nextRow, 1).Value = f
            master.Cells(nextRow, 2).Value = Now
            master.Cells(nextRow, 3).Value = arr(i, 1)
            master.Cells(nextRow, 4).Value = arr(i, 2)
            nextRow = nextRow + 1
        Next
        
        wb.Close SaveChanges:=False
        f = Dir()
    Loop
    Application.ScreenUpdating = True
End Sub
```

### 2.2 Build KPIs (PivotTable: Mean/StDev by Field and Source)

```vb
Option Explicit

Public Sub BuildKPIsPivot()
    Dim src As Worksheet: Set src = ThisWorkbook.Worksheets("MASTER")
    Dim dst As Worksheet
    On Error Resume Next
    Set dst = ThisWorkbook.Worksheets("KPIs")
    On Error GoTo 0
    If dst Is Nothing Then
        Set dst = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        dst.Name = "KPIs"
    Else
        dst.Cells.Clear
    End If
    
    Dim pc As PivotCache
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=src.UsedRange)
    
    Dim pt As PivotTable
    Set pt = pc.CreatePivotTable(TableDestination:=dst.Range("A3"), TableName:="ptKPIs")
    
    With pt
        .PivotFields("Field").Orientation = xlRowField
        .PivotFields("SourceFile").Orientation = xlColumnField
        .AddDataField .PivotFields("Value"), "Mean", xlAverage
        .AddDataField .PivotFields("Value"), "StDev", xlStDev
    End With
    
    dst.Range("A1").Value = "KPIs (Mean & StDev by Field and Source)"
End Sub
```

### 2.3 Resample a time series to uniform Δt (linear interpolation)

```vb
Option Explicit

' Inputs: Active sheet with A: time [s] (ascending), B: value
' Config!B2 = Δt [s]
Public Sub ResampleActiveSeries()
    Dim cfg As Worksheet: Set cfg = ThisWorkbook.Worksheets("Config")
    Dim dt As Double: dt = cfg.Range("B2").Value
    If dt <= 0# Then
        MsgBox "Set Config!B2 to a positive Δt (s).", vbExclamation
        Exit Sub
    End If
    
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim t() As Double, v() As Double, n As Long
    n = lastRow - 1
    If n < 2 Then
        MsgBox "Need at least two points.", vbExclamation
        Exit Sub
    End If
    
    Dim i As Long
    ReDim t(1 To n): ReDim v(1 To n)
    For i = 1 To n
        t(i) = ws.Cells(i + 1, "A").Value2
        v(i) = ws.Cells(i + 1, "B").Value2
    Next
    
    Dim tmin As Double, tmax As Double
    tmin = t(1): tmax = t(n)
    Dim m As Long: m = Fix((tmax - tmin) / dt) + 1
    
    Dim out() As Double: ReDim out(1 To m, 1 To 2)
    Dim tk As Double: tk = tmin
    
    Dim k As Long, j As Long: j = 1
    For k = 1 To m
        ' advance j such that t(j) <= tk <= t(j+1)
        Do While j < n And t(j + 1) < tk
            j = j + 1
        Loop
        Dim vk As Double
        If tk <= t(1) Then
            vk = v(1)
        ElseIf tk >= t(n) Then
            vk = v(n)
        Else
            Dim w As Double
            w = (tk - t(j)) / (t(j + 1) - t(j))
            vk = v(j) * (1 - w) + v(j + 1) * w
        End If
        out(k, 1) = tk
        out(k, 2) = vk
        tk = tk + dt
    Next
    
    Dim outWs As Worksheet
    Set outWs = ThisWorkbook.Worksheets.Add(After:=ws)
    outWs.Name = ws.Name & "_Resampled"
    outWs.Range("A1:B1").Value = Array("Time [s]", "Value")
    outWs.Range("A2").Resize(m, 2).Value = out
End Sub
```

***

## 3) Thermodynamics

### Summary

| Example                                       | Purpose                                 | Inputs                       | Outputs                     |
| --------------------------------------------- | --------------------------------------- | ---------------------------- | --------------------------- |
| Saturation pressure & boiling point (Antoine) | Quick Psat(T) and Tb(P)                 | T \[°C], A/B/C               | Psat \[kPa]; Boiling T at P |
| Heat exchanger (NTU‑ε method, counterflow)    | Compute duty and outlet temperatures    | `Thi, Tci, C_h, C_c, NTU`    | `Q`, `Tho`, `Tco`           |
| Moist air psychrometrics (simplified)         | Humidity ratio & enthalpy from T, RH, P | `Tdb[°C]`, `RH[-]`, `P[kPa]` | `W [kg/kg]`, `h [kJ/kg]`    |

> Approximations are standard engineering fits; document limits in your workbook.

### 3.1 Antoine equation utilities (Psat and Boiling Point)

```vb
Option Explicit

' Psat [kPa] from Antoine coefficients (A,B,C) with T in °C
Public Function Psat_Antoine_kPa(ByVal T_C As Double, ByVal A As Double, ByVal B As Double, ByVal C As Double) As Double
    ' Antoine: log10(P[bar or mmHg]) = A - B / (C + T)
    ' Here we expect kPa; use coefficients consistent with kPa
    ' If your A/B/C are for mmHg, convert: 1 mmHg = 0.133322 kPa
    Psat_Antoine_kPa = 10 ^ (A - B / (C + T_C))
End Function

' Solve Tb at pressure P using simple iteration (bisection)
Public Function BoilingPoint_Antoine(ByVal P_kPa As Double, ByVal A As Double, ByVal B As Double, ByVal C As Double, _
                                     Optional ByVal Tlow As Double = 0#, Optional ByVal Thigh As Double = 200#) As Double
    Dim lo As Double: lo = Tlow
    Dim hi As Double: hi = Thigh
    Dim mid As Double, Pmid As Double, iter As Long
    For iter = 1 To 60
        mid = 0.5 * (lo + hi)
        Pmid = Psat_Antoine_kPa(mid, A, B, C)
        If Pmid > P_kPa Then
            hi = mid
        Else
            lo = mid
        End If
    Next
    BoilingPoint_Antoine = 0.5 * (lo + hi)
End Function
```

> Add a small table of Antoine coefficients on a sheet (note the valid T range). For water, prefer fits appropriate to the temperature range you need.

### 3.2 Heat exchanger (counterflow) using NTU–ε

```vb
Option Explicit

' Effectiveness for counterflow: ε = (1 - exp(-NTU*(1-Cr))) / (1 - Cr*exp(-NTU*(1-Cr)))
Public Function HX_Eps_Counterflow(ByVal NTU As Double, ByVal Cr As Double) As Double
    If NTU <= 0# Then
        HX_Eps_Counterflow = 0#
        Exit Function
    End If
    If Abs(Cr - 1#) < 0.0001 Then
        HX_Eps_Counterflow = NTU / (1# + NTU) ' Cr≈1 special case
    Else
        HX_Eps_Counterflow = (1# - Exp(-NTU * (1# - Cr))) / (1# - Cr * Exp(-NTU * (1# - Cr)))
    End If
End Function

' Given Thi, Tci, C_h, C_c, NTU → compute Q, Tho, Tco (counterflow)
Public Sub HX_Compute_Counterflow()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("HX")
    ' Inputs in row 2: Thi (A), Tci (B), C_h (W/K) (C), C_c (W/K) (D), NTU (E)
    Dim Thi As Double, Tci As Double, Ch As Double, Cc As Double, NTU As Double
    Thi = ws.Range("A2").Value2
    Tci = ws.Range("B2").Value2
    Ch = ws.Range("C2").Value2
    Cc = ws.Range("D2").Value2
    NTU = ws.Range("E2").Value2
    
    Dim Cmin As Double, Cmax As Double, Cr As Double, eps As Double, Q As Double
    Cmin = Application.Min(Ch, Cc)
    Cmax = Application.Max(Ch, Cc)
    Cr = Cmin / Cmax
    eps = HX_Eps_Counterflow(NTU, Cr)
    Q = eps * Cmin * (Thi - Tci)
    
    Dim Tho As Double, Tco As Double
    If Ch = Cmin Then
        Tho = Thi - Q / Ch
        Tco = Tci + Q / Cc
    Else
        Tho = Thi - Q / Ch
        Tco = Tci + Q / Cc
    End If
    
    ws.Range("G1:J1").Value = Array("ε", "Q [W]", "Tho [°C]", "Tco [°C]")
    ws.Range("G2").Resize(1, 4).Value = Array(eps, Q, Tho, Tco)
End Sub
```

### 3.3 Simplified psychrometrics (Tdb, RH, P → W, h)

```vb
Option Explicit

' Tetens-like saturation pressure (kPa) for water, T in °C (0–50°C typical)
Private Function Pws_Tetens_kPa(ByVal T_C As Double) As Double
    Pws_Tetens_kPa = 0.61078 * Exp(17.2694 * T_C / (T_C + 237.29))
End Function

' Humidity ratio W [kg/kg dry air]
Public Function HumidityRatio_T_RH_P(ByVal Tdb_C As Double, ByVal RH As Double, ByVal P_kPa As Double) As Double
    Dim Pws As Double: Pws = Pws_Tetens_kPa(Tdb_C)
    Dim Pw As Double: Pw = RH * Pws
    HumidityRatio_T_RH_P = 0.62198 * Pw / (P_kPa - Pw)
End Function

' Moist air enthalpy (approx): h [kJ/kg dry air] = 1.006*T + W*(2501 + 1.86*T)
Public Function MoistAirEnthalpy_kJkg(ByVal Tdb_C As Double, ByVal W As Double) As Double
    MoistAirEnthalpy_kJkg = 1.006 * Tdb_C + W * (2501# + 1.86 * Tdb_C)
End Function

' Example table fill: A: Tdb [°C], B: RH [-], C: P [kPa] → D: W, E: h
Public Sub Psychro_Table()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Psychro")
    Dim r As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    ws.Range("D1:E1").Value = Array("W [kg/kg]", "h [kJ/kg]")
    For r = 2 To lastRow
        Dim T As Double, RH As Double, P As Double, W As Double
        T = ws.Cells(r, "A").Value2
        RH = ws.Cells(r, "B").Value2
        P = ws.Cells(r, "C").Value2
        W = HumidityRatio_T_RH_P(T, RH, P)
        ws.Cells(r, "D").Value = W
        ws.Cells(r, "E").Value = MoistAirEnthalpy_kJkg(T, W)
    Next
End Sub
```

***

## 4) Materials

### Summary

| Example                                               | Purpose                                         | Inputs                                         | Outputs                  |
| ----------------------------------------------------- | ----------------------------------------------- | ---------------------------------------------- | ------------------------ |
| Von Mises stress (3D state)                           | Safety check vs yield                           | Stress components: `σx, σy, σz, τxy, τyz, τzx` | `σvm`, pass/fail vs `σy` |
| Beam deflection & stress (simply supported, mid‑load) | Quick beam sanity checks                        | `F[N], L[m], b[m], h[m], E[Pa]`                | `δmax[m]`, `σmax[Pa]`    |
| Materials screening (FoM)                             | Rank candidates by E/ρ or σy/ρ with constraints | Table of materials                             | Ranked shortlist         |

### 4.1 Von Mises stress (general 3D)

```vb
Option Explicit

Public Function VonMises3D(ByVal sx As Double, ByVal sy As Double, ByVal sz As Double, _
                           ByVal txy As Double, ByVal tyz As Double, ByVal tzx As Double) As Double
    Dim term1 As Double, term2 As Double, term3 As Double
    term1 = (sx - sy) ^ 2 + (sy - sz) ^ 2 + (sz - sx) ^ 2
    term2 = 6# * (txy ^ 2 + tyz ^ 2 + tzx ^ 2)
    VonMises3D = Sqr(0.5 * (term1 + term2))
End Function

' Table utility: A: σx, B: σy, C: σz, D: τxy, E: τyz, F: τzx, G: σyld → H: σvm, I: Status
Public Sub VonMises_Table()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Stress")
    Dim r As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    ws.Range("H1:I1").Value = Array("σvm [Pa]", "Status")
    
    For r = 2 To lastRow
        Dim sx As Double, sy As Double, sz As Double, txy As Double, tyz As Double, tzx As Double, syld As Double, svm As Double
        sx = ws.Cells(r, "A").Value2
        sy = ws.Cells(r, "B").Value2
        sz = ws.Cells(r, "C").Value2
        txy = ws.Cells(r, "D").Value2
        tyz = ws.Cells(r, "E").Value2
        tzx = ws.Cells(r, "F").Value2
        syld = ws.Cells(r, "G").Value2
        
        svm = VonMises3D(sx, sy, sz, txy, tyz, tzx)
        ws.Cells(r, "H").Value = svm
        ws.Cells(r, "I").Value = IIf(svm <= syld, "PASS", "FAIL")
    Next
End Sub
```

### 4.2 Beam deflection & bending stress (simply supported, center point load)

```vb
Option Explicit

' For rectangular section: I = b*h^3/12, c = h/2
' δmax = F*L^3 / (48*E*I), σmax = Mmax*c/I with Mmax = F*L/4
Public Sub BeamSimpleMidLoad()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Beam")
    ' Inputs: A:F → F[N], L[m], b[m], h[m], E[Pa], Notes
    Dim r As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ws.Range("G1:H1").Value = Array("δmax [m]", "σmax [Pa]")
    
    For r = 2 To lastRow
        Dim F As Double, L As Double, b As Double, h As Double, E As Double
        F = ws.Cells(r, "A").Value2
        L = ws.Cells(r, "B").Value2
        b = ws.Cells(r, "C").Value2
        h = ws.Cells(r, "D").Value2
        E = ws.Cells(r, "E").Value2
        
        If F > 0# And L > 0# And b > 0# And h > 0# And E > 0# Then
            Dim I As Double, c As Double, delta As Double, sigma As Double
            I = b * h ^ 3 / 12#
            c = h / 2#
            delta = F * L ^ 3 / (48# * E * I)
            sigma = (F * L / 4#) * c / I
            ws.Cells(r, "G").Value = delta
            ws.Cells(r, "H").Value = sigma
        Else
            ws.Cells(r, "G").Value = ""
            ws.Cells(r, "H").Value = ""
        End If
    Next
End Sub
```

### 4.3 Materials screening by Figure‑of‑Merit (FoM)

```vb
Option Explicit

' Inputs sheet "Materials" with columns:
' A: Material, B: E [GPa], C: ρ [kg/m^3], D: σy [MPa], E: Tmax [°C], F: Cost [€/kg]
' Config on row 1: G1: FoMType ("E/rho" or "sigy/rho"), H1: Tmin [°C], I1: Tmax [°C], J1: MaxCost [€/kg]
' Outputs: K: FoM, L: Rank
Public Sub MaterialsScreen()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Materials")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim fomType As String: fomType = LCase$(ws.Range("G1").Value2)
    Dim Tmin As Double: Tmin = ws.Range("H1").Value2
    Dim Tmax As Double: Tmax = ws.Range("I1").Value2
    Dim Cmax As Double: Cmax = ws.Range("J1").Value2
    
    ws.Range("K1:L1").Value = Array("FoM", "Rank")
    
    Dim r As Long, count As Long: count = 0
    For r = 2 To lastRow
        Dim name As String, E_GPa As Double, rho As Double, sigy_MPa As Double, Tlim As Double, cost As Double, fom As Double
        name = ws.Cells(r, "A").Value
        E_GPa = ws.Cells(r, "B").Value2
        rho = ws.Cells(r, "C").Value2
        sigy_MPa = ws.Cells(r, "D").Value2
        Tlim = ws.Cells(r, "E").Value2
        cost = ws.Cells(r, "F").Value2
        
        ' Constraints
        If (Tlim >= Tmin And Tlim <= Tmax) And (cost <= Cmax) And (rho > 0#) Then
            Select Case fomType
                Case "e/rho": fom = (E_GPa * 1E9) / rho
                Case "sigy/rho", "σy/ρ": fom = (sigy_MPa * 1E6) / rho
                Case Else: fom = (E_GPa * 1E9) / rho
            End Select
            ws.Cells(r, "K").Value = fom
            count = count + 1
        Else
            ws.Cells(r, "K").Value = ""
        End If
    Next
    
    ' Rank FoM descending (simple approach)
    Dim rng As Range: Set rng = ws.Range("K2:K" & lastRow)
    Dim i As Long
    For r = 2 To lastRow
        If IsNumeric(ws.Cells(r, "K").Value) Then
            Dim rank As Long: rank = 1
            For i = 2 To lastRow
                If IsNumeric(ws.Cells(i, "K").Value) Then
                    If ws.Cells(i, "K").Value > ws.Cells(r, "K").Value Then rank = rank + 1
                End If
            Next
            ws.Cells(r, "L").Value = rank
        Else
            ws.Cells(r, "L").Value = ""
        End If
    Next
End Sub
```

***

## How to deploy these quickly

1.  Create an `.xlsm` workbook and add sheets named as referenced (`Ducts`, `Hydronic`, `Pumps`, `Config`, `MASTER`, `KPIs`, `Psychro`, `HX`, `Stress`, `Beam`, `Materials`) or adjust names in code.
2.  Paste each code block into its own standard module in the VBA editor (Alt+F11 → Insert → Module).
3.  Fill the input columns as specified in each summary table.
4.  Run macros (Alt+F8) and validate outputs.
5.  For repeatability, put **buttons** on each sheet linked to the corresponding macro and add a header row indicating units.

***

## Would you like these wrapped into a starter workbook?

If helpful, I can package these into a single `.xlsm` template with:

*   pre‑formatted tables,
*   buttons to run each macro,
*   a “Config” sheet (folder path, Δt, filters), and
*   a lightweight “Log” sheet that timestamps each run.

If you have specific fluid properties, duct roughness standards, motor lists, or materials catalogs your team uses, point me to those and I’ll align the defaults and unit tags accordingly.

# Error Handling

Absolutely. Below are **practical, reusable error‑handling patterns** you can drop into your workbook and **error‑hardened versions** of the earlier examples (HVAC hydraulics, test data consolidation, thermodynamics, materials). They demonstrate:

*   Centralized error logger (to a “Log” sheet)
*   Safe **Try/Catch/Finally**‑like structure (`On Error GoTo ErrH` + `CleanExit`)
*   Input validation and **domain checks** (engineering sanity bounds)
*   **Propagation** with `Err.Raise` and **context stacking**
*   **Retry** (e.g., file I/O) and **guarded toggles** (screen update, events, calc)
*   **Silent tests** with `On Error Resume Next` (only when justified)
*   Developer aids: `Debug.Assert`, `Stop`, and watchable `Err.Source`/`Err.Number`

> **Tip**: Keep error handling **consistent** across modules; copy the small **`modErrorUtils`** shown below and call its utilities everywhere.

***

## 0) Reusable Error Utilities (drop into `modErrorUtils`)

```vb
Option Explicit

' === Central logging to a "Log" sheet with timestamp ===
Public Sub LogMsg(ByVal level As String, ByVal where As String, ByVal msg As String, _
                  Optional ByVal errNum As Long = 0)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Log")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Log"
        ws.Range("A1:D1").Value = Array("Timestamp", "Level", "Where", "Message")
        ws.Range("E1").Value = "Err.Number"
    End If
    
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    ws.Cells(r, "A").Value = Now
    ws.Cells(r, "B").Value = UCase$(level)
    ws.Cells(r, "C").Value = where
    ws.Cells(r, "D").Value = msg
    If errNum <> 0 Then ws.Cells(r, "E").Value = errNum
End Sub

' === Guarded application speed toggles (always restore) ===
Public Sub SpeedStart(ByRef prevCalc As XlCalculation)
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

Public Sub SpeedEnd(ByVal prevCalc As XlCalculation)
    Application.EnableEvents = True
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True
End Sub

' === Validation helpers ===
Public Function MustBePositive(ByVal v As Double, ByVal fieldName As String) As Double
    If Not IsNumeric(v) Or v <= 0# Then
        Err.Raise vbObjectError + 1101, "Validation", fieldName & " must be positive."
    End If
    MustBePositive = v
End Function

Public Sub RequireSheet(ByVal sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1102, "Validation", "Missing sheet: " & sheetName
    End If
End Sub

' === Propagate error with extra context ===
Public Sub RethrowWithContext(ByVal context As String)
    Dim n As Long: n = Err.Number
    Dim d As String: d = Err.Description
    ' Preserve original number if non-zero, add context to description
    Err.Raise IIf(n <> 0, n, vbObjectError + 1999), context, context & " -> " & d
End Sub
```

***

## 1) Patterns You’ll Reuse Everywhere

### 1.1 Structured handler with “finally”

```vb
Sub Template_SafeBlock()
    On Error GoTo ErrH
    Dim prevCalc As XlCalculation
    SpeedStart prevCalc
    
    ' ... your logic ...
    
CleanExit:
    SpeedEnd prevCalc
    Exit Sub
ErrH:
    LogMsg "ERROR", "Template_SafeBlock", Err.Description, Err.Number
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Template"
    Resume CleanExit
End Sub
```

### 1.2 Guarded “probe” with `Resume Next` (safe only for *expected* failures)

```vb
Function SheetExists(ByVal name As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(name)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function
```

### 1.3 Retry pattern (e.g., transient file locks)

```vb
Function OpenWorkbookWithRetry(path As String, Optional maxRetries As Long = 3, Optional delayMs As Long = 300) As Workbook
    Dim i As Long
    For i = 1 To maxRetries
        On Error Resume Next
        Set OpenWorkbookWithRetry = Workbooks.Open(Filename:=path, ReadOnly:=True)
        If Err.Number = 0 Then Exit Function
        On Error GoTo 0
        DoEvents
        Application.Wait Now + (delayMs / 86400000#) ' ms → days
    Next
    Err.Raise vbObjectError + 1201, "OpenWorkbookWithRetry", "Failed to open after retries: " & path
End Function
```

### 1.4 Developer‑only checks

```vb
Sub DevChecks()
    Debug.Assert 2 + 2 = 4
    ' Force a break in debug sessions
    ' Stop
End Sub
```

***

## 2) HVAC Hydraulics — Error‑Hardened Examples

### 2.1 Duct Pressure Drop (with validation & context)

```vb
Option Explicit

Private Function FricHaaland(ByVal Re As Double, ByVal relR As Double) As Double
    If Re <= 0# Then
        FricHaaland = 0#
    Else
        Dim invSqrtF As Double
        invSqrtF = -1.8 * Log10((relR / 3.7) ^ 1.11 + 6.9 / Re)
        FricHaaland = 1# / (invSqrtF ^ 2)
    End If
End Function

Public Sub DuctPressureDropSegments_Safe()
    On Error GoTo ErrH
    Dim prevCalc As XlCalculation
    SpeedStart prevCalc
    
    RequireSheet "Ducts"
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Ducts")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Err.Raise vbObjectError + 1301, "HVAC", "No data rows in Ducts."

    ws.Range("H1:K1").Value = Array("Re", "f", "V [m/s]", "ΔP [Pa]")
    
    Dim r As Long
    For r = 2 To lastRow
        Dim L As Double, D As Double, eps As Double, Q As Double, rho As Double, mu As Double
        L = MustBePositive(ws.Cells(r, "B").Value2, "Length L")
        D = MustBePositive(ws.Cells(r, "C").Value2, "Diameter D")
        eps = ws.Cells(r, "D").Value2
        If eps < 0# Then Err.Raise vbObjectError + 1302, "HVAC", "Roughness ε cannot be negative."
        Q = ws.Cells(r, "E").Value2
        rho = MustBePositive(ws.Cells(r, "F").Value2, "Density ρ")
        mu = MustBePositive(ws.Cells(r, "G").Value2, "Viscosity μ")
        
        Dim A As Double, V As Double, Re As Double, relR As Double, f As Double, dP As Double
        A = WorksheetFunction.Pi() * D ^ 2 / 4#
        If A = 0# Then Err.Raise vbObjectError + 1303, "HVAC", "Zero area (check D)."
        
        V = IIf(Q <> 0#, Q / A, 0#)
        Re = IIf(mu <> 0#, rho * V * D / mu, 0#)
        relR = eps / D
        
        If Re > 0# And Re < 2300# Then
            f = 64# / Re
        Else
            f = FricHaaland(Re, relR)
        End If
        
        dP = f * (L / D) * (rho * V * V / 2#)
        
        ws.Cells(r, "H").Value = Re
        ws.Cells(r, "I").Value = f
        ws.Cells(r, "J").Value = V
        ws.Cells(r, "K").Value = dP
    Next
    
CleanExit:
    SpeedEnd prevCalc
    Exit Sub
ErrH:
    RethrowWithContext "DuctPressureDropSegments_Safe"
    LogMsg "ERROR", "DuctPressureDropSegments_Safe", Err.Description, Err.Number
    MsgBox "Duct calc failed: " & Err.Description, vbCritical
    Resume CleanExit
End Sub
```

### 2.2 Hydronic Balancing (guard numeric issues)

```vb
Public Sub HydronicBalanceRecommend_Safe()
    On Error GoTo ErrH
    Dim prevCalc As XlCalculation
    SpeedStart prevCalc
    
    RequireSheet "Hydronic"
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Hydronic")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Err.Raise vbObjectError + 1310, "Hydronic", "No data rows."
    
    ws.Range("E1:F1").Value = Array("Pos_new [%]", "Note")
    Dim r As Long
    For r = 2 To lastRow
        Dim Qd As Double, Qm As Double, posNow As Double, posNew As Double
        Qd = MustBePositive(ws.Cells(r, "B").Value2, "Q_design")
        Qm = MustBePositive(ws.Cells(r, "C").Value2, "Q_measured")
        posNow = MustBePositive(ws.Cells(r, "D").Value2, "Position_current")
        
        posNew = posNow * (Qd / Qm) ^ 2
        posNew = Application.Min(Application.Max(posNew, 5#), 100#)
        ws.Cells(r, "E").Value = posNew
        ws.Cells(r, "F").Value = IIf(Abs(Qd - Qm) / Qd < 0.05, "OK (<5%)", "Adjust")
    Next
    
CleanExit:
    SpeedEnd prevCalc
    Exit Sub
ErrH:
    LogMsg "ERROR", "HydronicBalanceRecommend_Safe", Err.Description, Err.Number
    MsgBox "Hydronic balancing failed: " & Err.Description, vbCritical
    Resume CleanExit
End Sub
```

### 2.3 Pump Power & Motor Selection (with bounds and messages)

```vb
Public Sub PumpPowerSelect_Safe()
    On Error GoTo ErrH
    Dim prevCalc As XlCalculation
    SpeedStart prevCalc
    
    RequireSheet "Pumps"
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Pumps")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Err.Raise vbObjectError + 1320, "Pumps", "No data rows."
    
    ws.Range("E1:F1").Value = Array("P_shaft [kW]", "Motor_kW (nearest)")
    Dim g As Double: g = 9.80665
    Dim motors As Variant
    motors = Array(0.37, 0.55, 0.75, 1.1, 1.5, 2.2, 3, 4, 5.5, 7.5, 11, 15, 18.5, 22, 30, 37, 45, 55, 75, 90, 110, 132)
    
    Dim r As Long
    For r = 2 To lastRow
        Dim Q As Double, H As Double, rho As Double, etaPct As Double, eta As Double
        Q = MustBePositive(ws.Cells(r, "A").Value2, "Q")
        H = MustBePositive(ws.Cells(r, "B").Value2, "H")
        rho = MustBePositive(ws.Cells(r, "C").Value2, "ρ")
        etaPct = MustBePositive(ws.Cells(r, "D").Value2, "η_pump [%]")
        eta = etaPct / 100#
        If eta <= 0# Or eta > 1# Then Err.Raise vbObjectError + 1321, "Pumps", "η_pump [%] must be (0,100]."
        
        Dim PkW As Double: PkW = rho * g * Q * H / eta / 1000#
        ws.Cells(r, "E").Value = PkW
        
        Dim i As Long, sel As Double: sel = motors(UBound(motors))
        For i = LBound(motors) To UBound(motors)
            If PkW <= motors(i) * 0.9 Then sel = motors(i): Exit For
        Next
        ws.Cells(r, "F").Value = sel
    Next
    
CleanExit:
    SpeedEnd prevCalc
    Exit Sub
ErrH:
    LogMsg "ERROR", "PumpPowerSelect_Safe", Err.Description, Err.Number
    MsgBox "Pump sizing failed: " & Err.Description, vbCritical
    Resume CleanExit
End Sub
```

***

## 3) Test Data Consolidation — Error‑Hardened Examples

### 3.1 Import CSV Folder → MASTER (with retry + summary)

```vb
Option Explicit

Public Sub ImportCSVFolderToMaster_Safe()
    On Error GoTo ErrH
    Dim prevCalc As XlCalculation
    SpeedStart prevCalc
    
    RequireSheet "Config"
    Dim cfg As Worksheet: Set cfg = ThisWorkbook.Worksheets("Config")
    Dim folder As String: folder = cfg.Range("B1").Value
    If Len(folder) = 0 Then Err.Raise vbObjectError + 1401, "Import", "Config!B1 must contain a folder path."
    If Right$(folder, 1) <> "\" And Right$(folder, 1) <> "/" Then folder = folder & Application.PathSeparator
    
    Dim master As Worksheet
    On Error Resume Next
    Set master = ThisWorkbook.Worksheets("MASTER")
    On Error GoTo 0
    If master Is Nothing Then
        Set master = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        master.Name = "MASTER"
    Else
        master.Cells.Clear
    End If
    master.Range("A1:D1").Value = Array("SourceFile", "Timestamp", "Field", "Value")
    
    Dim f As String: f = Dir(folder & "*.csv")
    If Len(f) = 0 Then
        LogMsg "WARN", "ImportCSVFolderToMaster_Safe", "No CSV files found in " & folder
        MsgBox "No CSV found in folder: " & folder, vbInformation
        GoTo CleanExit
    End If
    
    Dim nextRow As Long: nextRow = 2
    Dim countFiles As Long, countRows As Long
    
    Do While Len(f) > 0
        Dim wb As Workbook
        Set wb = OpenWorkbookWithRetry(folder & f, 3, 300)
        Dim ws As Worksheet: Set ws = wb.Worksheets(1)
        
        Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If lastRow >= 2 Then
            Dim rng As Range: Set rng = ws.Range("A1:B" & lastRow)
            Dim arr As Variant: arr = rng.Value
            Dim i As Long
            ' Basic header check
            If LCase$(CStr(arr(1, 1))) <> "field" Or LCase$(CStr(arr(1, 2))) <> "value" Then
                LogMsg "WARN", "Import", "Skipping file (unexpected headers): " & f
                wb.Close SaveChanges:=False
                f = Dir(): GoTo ContinueLoop
            End If
            
            For i = 2 To UBound(arr, 1)
                master.Cells(nextRow, 1).Value = f
                master.Cells(nextRow, 2).Value = Now
                master.Cells(nextRow, 3).Value = arr(i, 1)
                master.Cells(nextRow, 4).Value = arr(i, 2)
                nextRow = nextRow + 1
            Next
            countRows = countRows + (UBound(arr, 1) - 1)
        End If
        
        wb.Close SaveChanges:=False
        countFiles = countFiles + 1
ContinueLoop:
        f = Dir()
    Loop
    
    LogMsg "INFO", "ImportCSVFolderToMaster_Safe", "Imported files: " & countFiles & ", rows: " & countRows
    
CleanExit:
    SpeedEnd prevCalc
    Exit Sub
ErrH:
    LogMsg "ERROR", "ImportCSVFolderToMaster_Safe", Err.Description, Err.Number
    MsgBox "Import failed: " & Err.Description, vbCritical
    Resume CleanExit
End Sub
```

### 3.2 Build KPIs (Pivot) — handle caching errors

```vb
Public Sub BuildKPIsPivot_Safe()
    On Error GoTo ErrH
    RequireSheet "MASTER"
    Dim src As Worksheet: Set src = ThisWorkbook.Worksheets("MASTER")
    If src.UsedRange.Rows.Count < 2 Then Err.Raise vbObjectError + 1410, "KPIs", "MASTER is empty."
    
    Dim dst As Worksheet
    On Error Resume Next
    Set dst = ThisWorkbook.Worksheets("KPIs")
    On Error GoTo 0
    If dst Is Nothing Then
        Set dst = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        dst.Name = "KPIs"
    Else
        dst.Cells.Clear
    End If
    
    Dim pc As PivotCache
    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=src.UsedRange)
    
    Dim pt As PivotTable
    Set pt = pc.CreatePivotTable(TableDestination:=dst.Range("A3"), TableName:="ptKPIs")
    
    With pt
        .PivotFields("Field").Orientation = xlRowField
        .PivotFields("SourceFile").Orientation = xlColumnField
        .AddDataField .PivotFields("Value"), "Mean", xlAverage
        .AddDataField .PivotFields("Value"), "StDev", xlStDev
    End With
    dst.Range("A1").Value = "KPIs (Mean & StDev by Field and Source)"
    
    LogMsg "INFO", "BuildKPIsPivot_Safe", "Pivot built successfully.", 0
    Exit Sub
ErrH:
    LogMsg "ERROR", "BuildKPIsPivot_Safe", Err.Description, Err.Number
    MsgBox "Building KPIs failed: " & Err.Description, vbCritical
End Sub
```

***

## 4) Thermodynamics — Error‑Hardened Examples

### 4.1 Antoine Utilities (parameter sanity + bracket failure)

```vb
Option Explicit

Public Function Psat_Antoine_kPa_Safe(ByVal T_C As Double, ByVal A As Double, ByVal B As Double, ByVal C As Double) As Double
    On Error GoTo ErrH
    ' Basic sanity: avoid singularity at T = -C
    If Abs(T_C + C) < 0.0001 Then Err.Raise vbObjectError + 1501, "Antoine", "T near -C causes singularity."
    Psat_Antoine_kPa_Safe = 10 ^ (A - B / (C + T_C))
    Exit Function
ErrH:
    RethrowWithContext "Psat_Antoine_kPa_Safe"
End Function

Public Function BoilingPoint_Antoine_Safe(ByVal P_kPa As Double, ByVal A As Double, ByVal B As Double, ByVal C As Double, _
                                          Optional ByVal Tlow As Double = 0#, Optional ByVal Thigh As Double = 200#) As Double
    On Error GoTo ErrH
    Dim lo As Double: lo = Tlow
    Dim hi As Double: hi = Thigh
    If hi <= lo Then Err.Raise vbObjectError + 1502, "Antoine", "Invalid bracket [Tlow, Thigh]."
    
    Dim iter As Long, mid As Double, Pmid As Double
    For iter = 1 To 60
        mid = 0.5 * (lo + hi)
        Pmid = Psat_Antoine_kPa_Safe(mid, A, B, C)
        If Pmid > P_kPa Then
            hi = mid
        Else
            lo = mid
        End If
    Next
    BoilingPoint_Antoine_Safe = 0.5 * (lo + hi)
    Exit Function
ErrH:
    RethrowWithContext "BoilingPoint_Antoine_Safe"
End Function
```

### 4.2 Heat Exchanger NTU‑ε (guard Cr≈1 and invalid inputs)

```vb
Public Function HX_Eps_Counterflow_Safe(ByVal NTU As Double, ByVal Cr As Double) As Double
    On Error GoTo ErrH
    If NTU <= 0# Then Err.Raise vbObjectError + 1510, "HX", "NTU must be > 0."
    If Cr < 0# Or Cr > 1# Then Err.Raise vbObjectError + 1511, "HX", "Cr must be in [0,1]."
    If Abs(Cr - 1#) < 0.0001 Then
        HX_Eps_Counterflow_Safe = NTU / (1# + NTU)
    Else
        HX_Eps_Counterflow_Safe = (1# - Exp(-NTU * (1# - Cr))) / (1# - Cr * Exp(-NTU * (1# - Cr)))
    End If
    Exit Function
ErrH:
    RethrowWithContext "HX_Eps_Counterflow_Safe"
End Function

Public Sub HX_Compute_Counterflow_Safe()
    On Error GoTo ErrH
    RequireSheet "HX"
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("HX")
    Dim Thi As Double, Tci As Double, Ch As Double, Cc As Double, NTU As Double
    Thi = ws.Range("A2").Value2
    Tci = ws.Range("B2").Value2
    Ch = MustBePositive(ws.Range("C2").Value2, "C_h")
    Cc = MustBePositive(ws.Range("D2").Value2, "C_c")
    NTU = MustBePositive(ws.Range("E2").Value2, "NTU")
    
    Dim Cmin As Double, Cmax As Double, Cr As Double
    Cmin = Application.Min(Ch, Cc)
    Cmax = Application.Max(Ch, Cc)
    Cr = Cmin / Cmax
    
    Dim eps As Double: eps = HX_Eps_Counterflow_Safe(NTU, Cr)
    Dim Q As Double: Q = eps * Cmin * (Thi - Tci)
    Dim Tho As Double, Tco As Double
    Tho = Thi - Q / Ch
    Tco = Tci + Q / Cc
    
    ws.Range("G1:J1").Value = Array("ε", "Q [W]", "Tho [°C]", "Tco [°C]")
    ws.Range("G2:J2").Value = Array(eps, Q, Tho, Tco)
    
    Exit Sub
ErrH:
    LogMsg "ERROR", "HX_Compute_Counterflow_Safe", Err.Description, Err.Number
    MsgBox "Heat exchanger calc failed: " & Err.Description, vbCritical
End Sub
```

***

## 5) Materials — Error‑Hardened Examples

### 5.1 Von Mises (ensure numeric & pass/fail labeling)

```vb
Option Explicit

Public Function VonMises3D_Safe(ByVal sx As Double, ByVal sy As Double, ByVal sz As Double, _
                                ByVal txy As Double, ByVal tyz As Double, ByVal tzx As Double) As Double
    On Error GoTo ErrH
    Dim term1 As Double, term2 As Double
    term1 = (sx - sy) ^ 2 + (sy - sz) ^ 2 + (sz - sx) ^ 2
    term2 = 6# * (txy ^ 2 + tyz ^ 2 + tzx ^ 2)
    VonMises3D_Safe = Sqr(0.5 * (term1 + term2))
    Exit Function
ErrH:
    RethrowWithContext "VonMises3D_Safe"
End Function

Public Sub VonMises_Table_Safe()
    On Error GoTo ErrH
    RequireSheet "Stress"
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Stress")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Err.Raise vbObjectError + 1601, "Stress", "No data rows."
    
    ws.Range("H1:I1").Value = Array("σvm [Pa]", "Status")
    Dim r As Long
    For r = 2 To lastRow
        Dim sx As Double, sy As Double, sz As Double, txy As Double, tyz As Double, tzx As Double, syld As Double
        sx = ws.Cells(r, "A").Value2
        sy = ws.Cells(r, "B").Value2
        sz = ws.Cells(r, "C").Value2
        txy = ws.Cells(r, "D").Value2
        tyz = ws.Cells(r, "E").Value2
        tzx = ws.Cells(r, "F").Value2
        syld = MustBePositive(ws.Cells(r, "G").Value2, "Yield σy")
        
        Dim svm As Double: svm = VonMises3D_Safe(sx, sy, sz, txy, tyz, tzx)
        ws.Cells(r, "H").Value = svm
        ws.Cells(r, "I").Value = IIf(svm <= syld, "PASS", "FAIL")
    Next
    Exit Sub
ErrH:
    LogMsg "ERROR", "VonMises_Table_Safe", Err.Description, Err.Number
    MsgBox "Von Mises table failed: " & Err.Description, vbCritical
End Sub
```

### 5.2 Beam (validate geometry & modulus)

```vb
Public Sub BeamSimpleMidLoad_Safe()
    On Error GoTo ErrH
    RequireSheet "Beam"
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Beam")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Err.Raise vbObjectError + 1610, "Beam", "No data rows."
    
    ws.Range("G1:H1").Value = Array("δmax [m]", "σmax [Pa]")
    Dim r As Long
    For r = 2 To lastRow
        Dim F As Double, L As Double, b As Double, h As Double, E As Double
        F = MustBePositive(ws.Cells(r, "A").Value2, "Force F")
        L = MustBePositive(ws.Cells(r, "B").Value2, "Span L")
        b = MustBePositive(ws.Cells(r, "C").Value2, "Width b")
        h = MustBePositive(ws.Cells(r, "D").Value2, "Height h")
        E = MustBePositive(ws.Cells(r, "E").Value2, "Modulus E")
        
        Dim I As Double: I = b * h ^ 3 / 12#
        Dim c As Double: c = h / 2#
        Dim delta As Double: delta = F * L ^ 3 / (48# * E * I)
        Dim sigma As Double: sigma = (F * L / 4#) * c / I
        ws.Cells(r, "G").Value = delta
        ws.Cells(r, "H").Value = sigma
    Next
    Exit Sub
ErrH:
    LogMsg "ERROR", "BeamSimpleMidLoad_Safe", Err.Description, Err.Number
    MsgBox "Beam calc failed: " & Err.Description, vbCritical
End Sub
```

***

## 6) Additional Patterns You Might Need

### 6.1 Safe `.Find` usage (avoids hidden state issues)

```vb
Function FindExact(ByVal inRange As Range, ByVal what As String) As Range
    On Error GoTo ErrH
    Set FindExact = inRange.Find(What:=what, LookIn:=xlValues, LookAt:=xlWhole, _
                                 SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    Exit Function
ErrH:
    RethrowWithContext "FindExact"
End Function
```

### 6.2 Safe array write/read with dimension checks

```vb
Sub WriteArray2D(ByVal target As Range, ByRef arr As Variant)
    On Error GoTo ErrH
    If IsEmpty(arr) Then Err.Raise vbObjectError + 1701, "WriteArray2D", "Array is empty."
    If UBound(arr, 2) <> target.Columns.Count Or UBound(arr, 1) <> target.Rows.Count Then
        Err.Raise vbObjectError + 1702, "WriteArray2D", "Array size does not match target range."
    End If
    target.Value = arr
    Exit Sub
ErrH:
    RethrowWithContext "WriteArray2D"
End Sub
```

### 6.3 User‑facing consolidated messages after batch

```vb
Sub BatchWithSummary()
    On Error GoTo ErrH
    Dim okCount As Long, failCount As Long
    
    ' ... loop over items ...
    ' on each success: okCount = okCount + 1
    ' on recoverable failure: failCount = failCount + 1 : LogMsg "WARN", ...
    
CleanExit:
    MsgBox "Completed. OK=" & okCount & ", Failed=" & failCount, vbInformation
    Exit Sub
ErrH:
    LogMsg "ERROR", "BatchWithSummary", Err.Description, Err.Number
    MsgBox "Batch aborted: " & Err.Description, vbCritical
    Resume CleanExit
End Sub
```

***

## 7) How to integrate quickly

1.  **Add a “Log” sheet** (or let the logger create it).
2.  Paste `modErrorUtils` first; then replace your macros with the **`*_Safe`** variants above.
3.  Ensure required sheets exist (`Config`, `MASTER`, domain sheets).
4.  Run and inspect the **Log** on problems; messages include **where** and **Err.Number**.
5.  In development, enable **Immediate Window** (Ctrl+G) and use **breakpoints** and **watches**.
