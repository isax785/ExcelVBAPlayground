# Error Handling

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

Below are **practical, reusable error‑handling patterns** you can drop into your workbook and **error‑hardened versions** of the earlier examples (HVAC hydraulics, test data consolidation, thermodynamics, materials). They demonstrate:

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
