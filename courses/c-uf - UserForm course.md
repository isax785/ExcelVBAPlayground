# UserForm Course

- [UserForm Course](#userform-course)
  - [Course Overview](#course-overview)
    - [Learning Objectives](#learning-objectives)
  - [Module 1 — Environment \& GUI Basics](#module-1--environment--gui-basics)
  - [Module 2 — Controls, Properties \& Events](#module-2--controls-properties--events)
  - [Module 3 — Event‑Driven Programming Patterns](#module-3--eventdriven-programming-patterns)
  - [Module 4 — Validation, Error Handling \& UX Feedback](#module-4--validation-error-handling--ux-feedback)
  - [Module 5 — Data Binding with Worksheets](#module-5--data-binding-with-worksheets)
  - [Module 6 — Layout, Reusability \& Patterns](#module-6--layout-reusability--patterns)
  - [Module 7 — Engineering Examples with UserForms](#module-7--engineering-examples-with-userforms)
    - [Overview](#overview)
    - [7.1 `frmHydraulics` — HVAC Hydraulics Calculator](#71-frmhydraulics--hvac-hydraulics-calculator)
    - [7.2 `frmImportRuns` — Test Data Consolidation (Folder picker, preview, progress)](#72-frmimportruns--test-data-consolidation-folder-picker-preview-progress)
    - [7.3 `frmHX` — Thermodynamics: Counterflow NTU–ε](#73-frmhx--thermodynamics-counterflow-ntuε)
    - [7.4 `frmMaterials` — Materials Filter \& Ranking (FoM)](#74-frmmaterials--materials-filter--ranking-fom)
  - [Module 8 — Capstone: Multi‑Page Engineering Console](#module-8--capstone-multipage-engineering-console)
  - [Module 9 — Deployment \& QA](#module-9--deployment--qa)
  - [Suggested Schedule](#suggested-schedule)
  - [What you’ll take away](#what-youll-take-away)
  - [Quick Next Steps](#quick-next-steps)

---

Below is a **complete, engineering‑oriented course outline** for building **Excel VBA GUIs with UserForms**, including hands‑on labs and **four domain examples** (HVAC hydraulics, test data consolidation, thermodynamics, materials). It emphasizes robust patterns (validation, error handling, separation of concerns), fast workflows, and maintainability in engineering environments.

> **Note:** Code blocks are ready to paste into a `.xlsm` workbook. Keep a reusable `modErrorUtils` module (logger, speed toggles, validators) as introduced earlier; I reference it here.

***

## Course Overview

**Audience:** Engineers automating workflows who want interactive, reliable front‑ends in Excel.  
**Format:** 1–2 days, instructor‑led, with modular labs.  
**Outcomes:** Participants can design, code, and deploy professional UserForms that validate inputs, orchestrate calculations, write results to sheets, and handle errors gracefully.

### Learning Objectives

*   Design UserForms using the VBE **Toolbox**, property grid, and naming conventions.
*   Wire **events** and write clean controller logic (UI logic separated from calculations).
*   Validate inputs, **show clear feedback**, and log errors to a “Log” sheet.
*   Bind controls to worksheet data (Tables/ListObjects) for dynamic selection and display.
*   Build **engineering calculators** and data automation UIs with progress indication.
*   Package and deploy GUIs (buttons on sheets, add‑ins, digital signatures).

***

## Module 1 — Environment & GUI Basics

**Topics**

*   Enable Developer Tab; open VBE (Alt+F11).
*   Insert a **UserForm** (Insert → UserForm), place controls from **Toolbox**.
*   Windows: Project Explorer, Properties (F4), Code window.
*   **Naming conventions**: `frmHydraulics`, `txtDiameter`, `cmbFluid`, `btnCompute`, `lblResult`, `lstFiles`.
*   **Show/Unload** forms: `frmHydraulics.Show vbModeless` (non‑blocking) vs `vbModal` (blocking).
*   Organize code: `modUI_EntryPoints` (Show procedures), `modDomain_*` (calculations), `modErrorUtils` (logging/validation).

**Lab 1**

*   Create `frmHello` with a Label + “Close” button; wire `btnClose_Click` to `Unload Me`.

***

## Module 2 — Controls, Properties & Events

**Most‑used controls for engineering GUIs**

| Control                | Key Properties                        | Key Events                   | Typical Uses                                                  |
| ---------------------- | ------------------------------------- | ---------------------------- | ------------------------------------------------------------- |
| `TextBox`              | `Text`, `Value`, `Tag`, `Enabled`     | `Change`, `Exit`, `Validate` | Numeric inputs (diameter, flow, etc.)                         |
| `ComboBox`             | `RowSource`, `List`, `Value`, `Style` | `Change`, `DropButtonClick`  | Unit selection, fluid selection                               |
| `ListBox`              | `List`, `ListIndex`, `MultiSelect`    | `Change`, `DblClick`         | File lists, material catalogs                                 |
| `CheckBox`             | `Value`                               | `Click`                      | Options/toggles (include minor losses, use turbulence model…) |
| `OptionButton`         | `Value`, Grouping by frame            | `Click`                      | Mutually exclusive modes (Laminar/Turbulent Auto)             |
| `CommandButton`        | Caption                               | `Click`                      | Execute actions                                               |
| `Label`                | `Caption`, `BackColor`                | —                            | Display outputs/status                                        |
| `Frame`                | —                                     | —                            | Group related inputs visually                                 |
| `MultiPage`            | Pages collection                      | `Change`                     | Wizard/Tabbed UI                                              |
| `RefEdit`              | `Value`                               | —                            | Allow user to select a range                                  |
| `Image`                | `Picture`                             | —                            | Logos, diagrams                                               |
| `ScrollBar/SpinButton` | `Min/Max`, `Value`                    | `Change`                     | Fine‑tune numeric inputs                                      |
| `ProgressBar`\*        | —                                     | Custom (Frame+Label)         | Progress indication                                           |

\* No native ProgressBar control in standard toolbox; simulate with a **Frame + Label** (see later).

**Lab 2**

*   Build a form with `TextBox` + `ComboBox` inputs and a `Compute` button; on `Compute`, validate numbers and present a formatted result in a `Label`.

***

## Module 3 — Event‑Driven Programming Patterns

**Patterns**

*   `UserForm_Initialize`: populate lists, set defaults, configure fonts and units.
*   `Change` vs `Exit` vs `Validate`: validate on exit for numeric fields, compute on button click.
*   **Separation of concerns**: `Form` gathers inputs → calls `modDomain_*` functions → writes outputs back to form or sheet.
*   **Modeless forms** (`vbModeless`) for long‑running operations + `DoEvents`.
*   **Safe close**: `QueryClose` to confirm when tasks are running.

**Lab 3**

*   Populate a ComboBox with fluids from a “Fluids” sheet on `Initialize`, store selected fluid’s properties in hidden fields or a small in‑memory cache.

***

## Module 4 — Validation, Error Handling & UX Feedback

**Key techniques**

*   Use centralized **validators** (e.g., `MustBePositive`) and **logger** from `modErrorUtils`.
*   Guard numeric parsing (`CDbl`, `IsNumeric`) and show **inline feedback** (e.g., red border, status label).
*   Wrap `Compute` in `On Error GoTo ErrH` + `CleanExit`.
*   Use a **Status label** and **Progress bar** for long import tasks.
*   For recoverable issues, show **MessageBox** with actionable text; add details to **Log** sheet.

**Lab 4**

*   Add validation to all numeric inputs; color invalid TextBoxes (e.g., `BackColor = vbYellow`) and block compute until fixed.

***

## Module 5 — Data Binding with Worksheets

**Approaches**

*   **Populate by code**: read a Table (`ListObject`) into a Variant array and assign to `ComboBox.List`.
*   **RowSource** for simple static ranges (watch for volatility if rows change).
*   **Two‑way flow**: write outputs to named ranges or append to a data Table.
*   **RefEdit** to let users target ranges interactively.

**Lab 5**

*   Bind a materials list from `Materials` sheet to `ListBox`, allow filtering by constraints, update list dynamically.

***

## Module 6 — Layout, Reusability & Patterns

**Topics**

*   Consistent layout: alignment, tab order (`TabIndex`), accelerator keys (`&Compute`).
*   **MVP‑lite**: Form as View, a small Presenter module orchestrates, domain module computes.
*   Reusable dialogs: folder picker, “are you sure?”, yes/no prompts.
*   **Progress dialog** pattern with cancel support.
*   Internationalization basics (units/labels).

**Lab 6**

*   Create a reusable **Folder Picker** and a **Progress UI**; use them from two different forms.

***

## Module 7 — Engineering Examples with UserForms

### Overview

| Form            | Domain                  | Purpose                                                       | Sheets Used                               |
| --------------- | ----------------------- | ------------------------------------------------------------- | ----------------------------------------- |
| `frmHydraulics` | HVAC Hydraulics         | Darcy–Weisbach calculator with Swamee–Jain/Haaland            | `Hydraulics` (results), optional `Fluids` |
| `frmImportRuns` | Test Data Consolidation | Select folder, preview CSVs, import with progress & logging   | `Config`, `MASTER`, `Log`                 |
| `frmHX`         | Thermodynamics          | Counterflow HX NTU‑ε calculator with outputs                  | `HX`                                      |
| `frmMaterials`  | Materials               | Filter materials by constraints, rank by FoM, write shortlist | `Materials`                               |

> Each example shows **Initialize**, **Compute/Run**, **Validation**, and **Write‑back**.

***

### 7.1 `frmHydraulics` — HVAC Hydraulics Calculator

**Controls (suggested)**

*   TextBoxes: `txtRho`, `txtMu`, `txtQ`, `txtD`, `txtL`, `txtEps`
*   OptionButtons (in a Frame): `optLaminarAuto`, `optTurbulentAuto` (or a `ComboBox` for method)
*   Buttons: `btnCompute`, `btnAppendToSheet`, `btnClose`
*   Labels: `lblRe`, `lblf`, `lblV`, `lbldP`, `lblStatus`

**UserForm Code (core parts)**

```vba
Option Explicit

Private Sub UserForm_Initialize()
    On Error GoTo ErrH
    ' Defaults
    txtRho.Text = "1000"      ' kg/m^3
    txtMu.Text = "0.001"      ' Pa·s
    txtQ.Text = "0.002"       ' m^3/s
    txtD.Text = "0.05"        ' m
    txtL.Text = "10"          ' m
    txtEps.Text = "0.0001"    ' m
    optTurbulentAuto.Value = True
    lblStatus.Caption = "Ready."
    Exit Sub
ErrH:
    MsgBox "Init failed: " & Err.Description, vbCritical
End Sub

Private Function ParsePositive(ByVal tb As MSForms.TextBox, ByVal fieldName As String) As Double
    If Not IsNumeric(tb.Text) Then
        Err.Raise vbObjectError + 2001, "HydraulicsUI", fieldName & " must be numeric."
    End If
    Dim v As Double: v = CDbl(tb.Text)
    If v <= 0# Then Err.Raise vbObjectError + 2002, "HydraulicsUI", fieldName & " must be > 0."
    ParsePositive = v
End Function

Private Function FricHaaland(ByVal Re As Double, ByVal relR As Double) As Double
    If Re <= 0# Then FricHaaland = 0#: Exit Function
    Dim invSqrtF As Double
    invSqrtF = -1.8 * Log10((relR / 3.7) ^ 1.11 + 6.9 / Re)
    FricHaaland = 1# / (invSqrtF ^ 2)
End Function

Private Sub btnCompute_Click()
    On Error GoTo ErrH
    Dim rho As Double, mu As Double, Q As Double, D As Double, L As Double, eps As Double
    rho = ParsePositive(txtRho, "Density ρ")
    mu = ParsePositive(txtMu, "Viscosity μ")
    Q = CDbl(txtQ.Text) ' allow zero flow for check
    D = ParsePositive(txtD, "Diameter D")
    L = ParsePositive(txtL, "Length L")
    eps = CDbl(txtEps.Text)
    If eps < 0# Then Err.Raise vbObjectError + 2003, "HydraulicsUI", "Roughness ε cannot be negative."

    Dim A As Double, V As Double, Re As Double, relR As Double, f As Double, dP As Double
    A = WorksheetFunction.Pi() * D ^ 2 / 4#
    If A = 0# Then Err.Raise vbObjectError + 2004, "HydraulicsUI", "Zero area."
    V = IIf(Q <> 0#, Q / A, 0#)
    Re = IIf(mu <> 0#, rho * V * D / mu, 0#)
    relR = eps / D
    If Re > 0# And Re < 2300# Then
        f = 64# / Re
    Else
        f = FricHaaland(Re, relR)
    End If
    dP = f * (L / D) * (rho * V * V / 2#)

    lblRe.Caption = Format(Re, "0.0")
    lblf.Caption = Format(f, "0.0000")
    lblV.Caption = Format(V, "0.000")
    lbldP.Caption = Format(dP, "0.0") & " Pa"
    lblStatus.Caption = "Computed."
    Exit Sub
ErrH:
    lblStatus.Caption = "Error."
    MsgBox "Compute failed: " & Err.Description, vbCritical
End Sub

Private Sub btnAppendToSheet_Click()
    On Error GoTo ErrH
    ' Assumes labels already set by Compute
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Hydraulics")
    If ws.Range("A1").Value = "" Then
        ws.Range("A1:F1").Value = Array("ρ", "μ", "Q", "D", "L", "ε")
        ws.Range("G1:K1").Value = Array("Re", "f", "V", "ΔP [Pa]", "Stamp")
    End If
    Dim r As Long: r = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    ws.Cells(r, "A").Value = CDbl(txtRho.Text)
    ws.Cells(r, "B").Value = CDbl(txtMu.Text)
    ws.Cells(r, "C").Value = CDbl(txtQ.Text)
    ws.Cells(r, "D").Value = CDbl(txtD.Text)
    ws.Cells(r, "E").Value = CDbl(txtL.Text)
    ws.Cells(r, "F").Value = CDbl(txtEps.Text)
    ws.Cells(r, "G").Value = Val(lblRe.Caption)
    ws.Cells(r, "H").Value = Val(lblf.Caption)
    ws.Cells(r, "I").Value = Val(lblV.Caption)
    ws.Cells(r, "J").Value = Val(Replace(lbldP.Caption, " Pa", ""))
    ws.Cells(r, "K").Value = Now
    lblStatus.Caption = "Appended row " & r & "."
    Exit Sub
ErrH:
    MsgBox "Append failed: " & Err.Description, vbCritical
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
```

**Entry point (standard module)**

```vba
Option Explicit
Public Sub ShowHydraulicsUI()
    frmHydraulics.Show vbModeless  ' or vbModal if preferred
End Sub
```

***

### 7.2 `frmImportRuns` — Test Data Consolidation (Folder picker, preview, progress)

**Controls (suggested)**

*   TextBox: `txtFolder`
*   Button: `btnBrowse`, `btnScan`, `btnImport`, `btnClose`
*   ListBox: `lstFiles` (multiselect to choose subset)
*   Frame + Label: `fraProgress` + `lblProgress` to simulate a progress bar
*   Label: `lblStatus`

**UserForm Code (core parts)**

```vba
Option Explicit

Private Sub UserForm_Initialize()
    txtFolder.Text = ""
    lstFiles.Clear
    SetupProgress 0
    lblStatus.Caption = "Select a folder and Scan."
End Sub

Private Sub btnBrowse_Click()
    On Error GoTo ErrH
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "Select data folder"
        If .Show = -1 Then
            txtFolder.Text = .SelectedItems(1) & Application.PathSeparator
        End If
    End With
    Exit Sub
ErrH:
    MsgBox "Folder picker failed: " & Err.Description, vbCritical
End Sub

Private Sub btnScan_Click()
    On Error GoTo ErrH
    lstFiles.Clear
    Dim folder As String: folder = txtFolder.Text
    If Len(folder) = 0 Then Err.Raise vbObjectError + 2101, "ImportUI", "Folder not set."
    Dim f As String: f = Dir(folder & "*.csv")
    Do While Len(f) > 0
        lstFiles.AddItem f
        f = Dir()
    Loop
    lblStatus.Caption = lstFiles.ListCount & " file(s) found."
    Exit Sub
ErrH:
    MsgBox "Scan failed: " & Err.Description, vbCritical
End Sub

Private Sub btnImport_Click()
    On Error GoTo ErrH
    If lstFiles.ListCount = 0 Then Err.Raise vbObjectError + 2102, "ImportUI", "No files to import."
    Dim folder As String: folder = txtFolder.Text
    Dim ws As Worksheet, master As Worksheet
    Set master = EnsureSheet("MASTER")
    If master.Range("A1").Value = "" Then
        master.Range("A1:D1").Value = Array("SourceFile", "Timestamp", "Field", "Value")
    End If
    Dim nextRow As Long: nextRow = master.Cells(master.Rows.Count, "A").End(xlUp).Row + 1
    
    Dim selCount As Long: selCount = GetSelectedCount(lstFiles)
    Dim i As Long, processed As Long
    If selCount = 0 Then selCount = lstFiles.ListCount ' import all if nothing selected
    SetupProgress 0
    For i = 0 To lstFiles.ListCount - 1
        If lstFiles.Selected(i) Or GetSelectedCount(lstFiles) = 0 Then
            ImportOneCSV folder & lstFiles.List(i), master, nextRow
            processed = processed + 1
            UpdateProgress processed / selCount
            DoEvents
        End If
    Next
    lblStatus.Caption = "Import completed: " & (processed) & " file(s)."
    Exit Sub
ErrH:
    MsgBox "Import failed: " & Err.Description, vbCritical
End Sub

' Helpers

Private Function EnsureSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        EnsureSheet.Name = name
    End If
End Function

Private Function GetSelectedCount(lst As MSForms.ListBox) As Long
    Dim i As Long
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then GetSelectedCount = GetSelectedCount + 1
    Next
End Function

Private Sub ImportOneCSV(ByVal path As String, ByRef master As Worksheet, ByRef nextRow As Long)
    On Error GoTo ErrH
    Dim wb As Workbook
    Set wb = Workbooks.Open(Filename:=path, ReadOnly:=True)
    Dim ws As Worksheet: Set ws = wb.Worksheets(1)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow >= 2 Then
        Dim arr As Variant: arr = ws.Range("A1:B" & lastRow).Value
        ' Header sanity
        If LCase$(CStr(arr(1, 1))) <> "field" Or LCase$(CStr(arr(1, 2))) <> "value" Then
            wb.Close SaveChanges:=False
            Err.Raise vbObjectError + 2103, "ImportUI", "Unexpected headers in " & Dir(path)
        End If
        Dim i As Long
        For i = 2 To UBound(arr, 1)
            master.Cells(nextRow, 1).Value = Dir(path)
            master.Cells(nextRow, 2).Value = Now
            master.Cells(nextRow, 3).Value = arr(i, 1)
            master.Cells(nextRow, 4).Value = arr(i, 2)
            nextRow = nextRow + 1
        Next
    End If
    wb.Close SaveChanges:=False
    Exit Sub
ErrH:
    ' Log warning and continue with next file
    ' If you included modErrorUtils, call: LogMsg "WARN", "ImportOneCSV", Err.Description
    MsgBox "Skipping file due to error: " & path & vbCrLf & Err.Description, vbExclamation
End Sub

' Progress bar: Frame (fraProgress) contains Label (lblProgress) docked left
Private Sub SetupProgress(ByVal frac As Double)
    lblProgress.Width = fraProgress.InsideWidth * frac
    lblProgress.Caption = Format(frac, "0%")
End Sub

Private Sub UpdateProgress(ByVal frac As Double)
    If frac < 0 Then frac = 0
    If frac > 1 Then frac = 1
    SetupProgress frac
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
```

**Entry point**

```vba
Option Explicit
Public Sub ShowImportRunsUI()
    frmImportRuns.Show vbModeless
End Sub
```

***

### 7.3 `frmHX` — Thermodynamics: Counterflow NTU–ε

**Controls**

*   TextBoxes: `txtThi`, `txtTci`, `txtCh`, `txtCc`, `txtNTU`
*   Button: `btnCompute`, `btnWrite`, `btnClose`
*   Labels: `lblEps`, `lblQ`, `lblTho`, `lblTco`, `lblStatus`

**UserForm Code (core parts)**

```vba
Option Explicit

Private Sub UserForm_Initialize()
    txtThi.Text = "80"   ' °C
    txtTci.Text = "20"   ' °C
    txtCh.Text = "2500"  ' W/K
    txtCc.Text = "3000"  ' W/K
    txtNTU.Text = "1.5"
    lblStatus.Caption = "Ready."
End Sub

Private Function ParsePos(tb As MSForms.TextBox, name As String) As Double
    If Not IsNumeric(tb.Text) Then Err.Raise vbObjectError + 2201, "HXUI", name & " must be numeric."
    ParsePos = CDbl(tb.Text)
    If ParsePos <= 0# Then Err.Raise vbObjectError + 2202, "HXUI", name & " must be > 0."
End Function

Private Function HX_Eps_Counterflow(ByVal NTU As Double, ByVal Cr As Double) As Double
    If Abs(Cr - 1#) < 0.0001 Then
        HX_Eps_Counterflow = NTU / (1# + NTU)
    Else
        HX_Eps_Counterflow = (1# - Exp(-NTU * (1# - Cr))) / (1# - Cr * Exp(-NTU * (1# - Cr)))
    End If
End Function

Private Sub btnCompute_Click()
    On Error GoTo ErrH
    Dim Thi As Double, Tci As Double, Ch As Double, Cc As Double, NTU As Double
    Thi = CDbl(txtThi.Text)  ' allow negative if needed (°C)
    Tci = CDbl(txtTci.Text)
    Ch = ParsePos(txtCh, "C_h")
    Cc = ParsePos(txtCc, "C_c")
    NTU = ParsePos(txtNTU, "NTU")
    
    Dim Cmin As Double, Cmax As Double, Cr As Double, eps As Double
    Cmin = Application.Min(Ch, Cc)
    Cmax = Application.Max(Ch, Cc)
    Cr = Cmin / Cmax
    If Cr < 0# Or Cr > 1# Then Err.Raise vbObjectError + 2203, "HXUI", "Cr must be in [0,1]."
    
    eps = HX_Eps_Counterflow(NTU, Cr)
    Dim Q As Double: Q = eps * Cmin * (Thi - Tci)
    Dim Tho As Double: Tho = Thi - Q / Ch
    Dim Tco As Double: Tco = Tci + Q / Cc
    
    lblEps.Caption = Format(eps, "0.000")
    lblQ.Caption = Format(Q, "0.0") & " W"
    lblTho.Caption = Format(Tho, "0.0") & " °C"
    lblTco.Caption = Format(Tco, "0.0") & " °C"
    lblStatus.Caption = "Computed."
    Exit Sub
ErrH:
    lblStatus.Caption = "Error."
    MsgBox "HX compute failed: " & Err.Description, vbCritical
End Sub

Private Sub btnWrite_Click()
    On Error GoTo ErrH
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("HX")
    If ws.Range("A1").Value = "" Then
        ws.Range("A1:E1").Value = Array("Thi", "Tci", "C_h", "C_c", "NTU")
        ws.Range("G1:J1").Value = Array("ε", "Q [W]", "Tho [°C]", "Tco [°C]")
    End If
    Dim r As Long: r = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    ws.Cells(r, "A").Value = CDbl(txtThi.Text)
    ws.Cells(r, "B").Value = CDbl(txtTci.Text)
    ws.Cells(r, "C").Value = CDbl(txtCh.Text)
    ws.Cells(r, "D").Value = CDbl(txtCc.Text)
    ws.Cells(r, "E").Value = CDbl(txtNTU.Text)
    ws.Cells(r, "G").Value = Val(lblEps.Caption)
    ws.Cells(r, "H").Value = Val(Replace(lblQ.Caption, " W", ""))
    ws.Cells(r, "I").Value = Val(Replace(lblTho.Caption, " °C", ""))
    ws.Cells(r, "J").Value = Val(Replace(lblTco.Caption, " °C", ""))
    lblStatus.Caption = "Written row " & r & "."
    Exit Sub
ErrH:
    MsgBox "Write failed: " & Err.Description, vbCritical
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
```

**Entry point**

```vba
Option Explicit
Public Sub ShowHXUI()
    frmHX.Show vbModal
End Sub
```

***

### 7.4 `frmMaterials` — Materials Filter & Ranking (FoM)

**Controls**

*   TextBoxes: `txtTmin`, `txtTmax`, `txtMaxCost`
*   ComboBox: `cmbFoM` (e.g., `E/ρ`, `σy/ρ`)
*   ListBox: `lstCandidates` (columns: Name, E, ρ, σy, Tmax, Cost, FoM)
*   Buttons: `btnFilter`, `btnWriteShortlist`, `btnClose`
*   Label: `lblStatus`

**UserForm Code (core parts)**

```vba
Option Explicit
Private materials As Variant  ' cached table
Private headers As Variant

Private Sub UserForm_Initialize()
    cmbFoM.Clear
    cmbFoM.AddItem "E/rho"
    cmbFoM.AddItem "sigy/rho"
    cmbFoM.ListIndex = 0
    txtTmin.Text = "0"
    txtTmax.Text = "200"
    txtMaxCost.Text = "100"
    LoadMaterials
End Sub

Private Sub LoadMaterials()
    On Error GoTo ErrH
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Materials")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Err.Raise vbObjectError + 2301, "MaterialsUI", "Materials sheet empty."
    headers = ws.Range("A1:F1").Value
    materials = ws.Range("A2:F" & lastRow).Value ' 2D array
    lblStatus.Caption = (lastRow - 1) & " materials loaded."
    Exit Sub
ErrH:
    MsgBox "Load failed: " & Err.Description, vbCritical
End Sub

Private Sub btnFilter_Click()
    On Error GoTo ErrH
    Dim Tmin As Double, Tmax As Double, Cmax As Double
    Tmin = CDbl(txtTmin.Text)
    Tmax = CDbl(txtTmax.Text)
    Cmax = CDbl(txtMaxCost.Text)
    Dim mode As String: mode = LCase$(cmbFoM.Text)
    
    lstCandidates.Clear
    lstCandidates.ColumnCount = 7
    lstCandidates.ColumnWidths = "100 pt;60 pt;60 pt;60 pt;60 pt;60 pt;80 pt"
    
    Dim r As Long, added As Long
    For r = LBound(materials, 1) To UBound(materials, 1)
        Dim name As String, E_GPa As Double, rho As Double, sigy_MPa As Double, TmaxRow As Double, cost As Double
        name = materials(r, 1)
        E_GPa = Val(materials(r, 2))
        rho = Val(materials(r, 3))
        sigy_MPa = Val(materials(r, 4))
        TmaxRow = Val(materials(r, 5))
        cost = Val(materials(r, 6))
        If rho > 0# And cost <= Cmax And TmaxRow >= Tmin And TmaxRow <= Tmax Then
            Dim fom As Double
            If mode = "e/rho" Then
                fom = (E_GPa * 1E9) / rho
            Else
                fom = (sigy_MPa * 1E6) / rho
            End If
            lstCandidates.AddItem
            lstCandidates.List(added, 0) = name
            lstCandidates.List(added, 1) = E_GPa
            lstCandidates.List(added, 2) = rho
            lstCandidates.List(added, 3) = sigy_MPa
            lstCandidates.List(added, 4) = TmaxRow
            lstCandidates.List(added, 5) = cost
            lstCandidates.List(added, 6) = Format(fom, "0.00E+00")
            added = added + 1
        End If
    Next
    lblStatus.Caption = added & " candidate(s)."
    Exit Sub
ErrH:
    MsgBox "Filter failed: " & Err.Description, vbCritical
End Sub

Private Sub btnWriteShortlist_Click()
    On Error GoTo ErrH
    If lstCandidates.ListCount = 0 Then Err.Raise vbObjectError + 2302, "MaterialsUI", "No candidates to write."
    Dim ws As Worksheet: Set ws = EnsureSheet("Shortlist")
    ws.Cells.Clear
    ws.Range("A1:G1").Value = Array("Material", "E [GPa]", "ρ [kg/m3]", "σy [MPa]", "Tmax [°C]", "Cost [€/kg]", "FoM")
    Dim r As Long
    For r = 0 To lstCandidates.ListCount - 1
        ws.Cells(r + 2, 1).Resize(1, 7).Value = lstRowToArray(lstCandidates, r)
    Next
    lblStatus.Caption = "Shortlist written: " & lstCandidates.ListCount & " rows."
    Exit Sub
ErrH:
    MsgBox "Write shortlist failed: " & Err.Description, vbCritical
End Sub

Private Function EnsureSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        EnsureSheet.Name = name
    End If
End Function

Private Function lstRowToArray(lst As MSForms.ListBox, ByVal idx As Long) As Variant
    Dim tmp(1 To 1, 1 To 7) As Variant
    Dim c As Long
    For c = 0 To 6
        tmp(1, c + 1) = lst.List(idx, c)
    Next
    lstRowToArray = tmp
End Function

Private Sub btnClose_Click()
    Unload Me
End Sub
```

**Entry point**

```vba
Option Explicit
Public Sub ShowMaterialsUI()
    frmMaterials.Show vbModal
End Sub
```

***

## Module 8 — Capstone: Multi‑Page Engineering Console

**Brief:** Build `frmEngineeringConsole` with **four tabs** (Hydraulics, Import, HX, Materials). Reuse the earlier logic by moving computation and import routines into **standard modules**, and let each tab call those routines. Add a **global status bar** and **log viewer** (button opens the “Log” sheet).

**Deliverables**

*   A single `.xlsm` with: `frmEngineeringConsole`, domain forms (optional), `modDomain_*`, `modUI_EntryPoints`, `modErrorUtils`, and example sheets (`Hydraulics`, `MASTER`, `HX`, `Materials`, `Log`).
*   Buttons on a **Home** sheet to open each UI.

***

## Module 9 — Deployment & QA

**Topics**

*   Add buttons on worksheets (Insert → Shape → Assign Macro `Show…UI`).
*   Macro security: sign the VBA project; set Trust Center appropriately.
*   Versioning: keep form code in **exported `.frm`/`.bas` files** for source control.
*   Performance: prefer **modeless** for long runs; update UI via `DoEvents`; batch Excel object calls; write arrays.
*   32/64‑bit: if calling WinAPI (not required here), use `PtrSafe` declares.
*   Documentation: a “README” sheet and tooltips (`ControlTipText`) in forms.

**QA Checklist**

*   Validate **all** numeric inputs; block compute on invalid fields.
*   Handle **empty datasets** gracefully (no crash, informative message).
*   Log to “Log” sheet: start/end timestamps, counts, warnings, errors.
*   Unit test domain functions (simple assertion procedures).

***

## Suggested Schedule

*   **Half‑day:** Modules 1–4 + one example form
*   **One day:** Modules 1–7 + labs
*   **Two days:** All modules + Capstone + reviews

***

## What you’ll take away

*   A **template UserForm** with validation, status, and progress.
*   Reusable entry points and domain modules ready for your team’s workflows.
*   Four engineering GUI examples adaptable to your standards and units.

***

## Quick Next Steps

*   Do you want the **multi‑page console** or separate forms per domain?
*   Should we target **modeless** operation (operators keep working in Excel during long imports)?
*   Share any **standard property tables** (fluids, materials) you use; I’ll wire them to the ComboBox/ListBox defaults.
