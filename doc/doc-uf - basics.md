# USER FORM - Basics

- [USER FORM - Basics](#user-form---basics)
- [What a UserForm Is (and Why It Matters)\*\*](#what-a-userform-is-and-why-it-matters)
  - [CREATE USER FORM](#create-user-form)
    - [INITIALIZATION](#initialization)
    - [CREATE A LIST FROM A RANGE IN A WORKSHEET](#create-a-list-from-a-range-in-a-worksheet)
    - [COMBO BOX](#combo-box)
    - [COMMAND BUTTON](#command-button)
  - [HOW TO CALL A USER FORM](#how-to-call-a-user-form)
    - [CALL FROM A COMMAND IN THE EXCEL FILE](#call-from-a-command-in-the-excel-file)
  - [OPEN USER FORM WHEN OPEN THE EXCEL FILE](#open-user-form-when-open-the-excel-file)
  - [WHAT TO DO WHEN CLOSE A USER FORM](#what-to-do-when-close-a-user-form)
  - [MAKE THE EXCEL APPLICATION VISIBLE BY A COMMAND BUTTON AND PSW](#make-the-excel-application-visible-by-a-command-button-and-psw)

---

# What a UserForm Is (and Why It Matters)**

A **UserForm** is a custom dialog window in Excel VBA that allows you to build interactive interfaces—data entry forms, dashboards, wizards, configuration panels, and more.

**Why UserForms are powerful**
- They enforce **structured data entry**
- They improve **user experience** and reduce errors
- They allow **custom workflows** beyond Excel’s built‑in dialogs
- They support **event‑driven programming**


## CREATE USER FORM

1. Open **VBA Editor** (ALT + F11)
2. Insert → UserForm
3. Add controls from the Toolbox
   1. Double click on the controls in order to program them.

### INITIALIZATION

Actions to do when the UserForm is initialized.

```vb
    Private Sub UserForm_Initialize()
    	...
    End Sub
```
### CREATE A LIST FROM A RANGE IN A WORKSHEET

If range in column:

```vb
    Me.ComboBox1.List = Range("A1:A6").Value  
```
If range in row:

```vb
    Me.ComboBox1.List = Application.Transpose(Range("A1:G1").Value)  
```

Other methods:

```vb
    Me.ComboBox1.List = Range(Selection, Selection.End(xlDown)).Value  
```

If merged cells in the range:

```vb
    Dim aCell As Range
    ...
    
    For Each aCell In Selection
        If aCell.Value <> "" Then
            Me.ComboBox1.AddItem aCell.Value
        End If
    Next
```
### COMBO BOX   

```vb
    Private Sub ComboBox1_Change()
    ...
    End Sub
```
### COMMAND BUTTON  

```vb
    Private Sub CommandButton1_Click()
    ...
    End Sub
```

## HOW TO CALL A USER FORM

### CALL FROM A COMMAND IN THE EXCEL FILE

In a worksheet, "Develop" --> "Insert" --> ActiveX Controls --> Command Button

Double click on the Button.

```vb
    Private Sub StartConversion_Click()
    	UserForm1.Show
    End Sub
```
## OPEN USER FORM WHEN OPEN THE EXCEL FILE 

Also, to hide the excel application.

```vb
    Private Sub Workbook_open()
    	Application.Visible = False
    	UserForm1.Show
    End Sub
```
## WHAT TO DO WHEN CLOSE A USER FORM

Make the excel application visible, close the workbook, ...

Example: close the workbook if the excel application is not visible (general case of generic user, not administrator).

```vb
    Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
        If Application.Visible = False Then
            Workbooks.Close
        End If
    End Sub
```

## MAKE THE EXCEL APPLICATION VISIBLE BY A COMMAND BUTTON AND PSW

```vb
    Private Sub CommandButton2_Click()
    Dim inp As String
    psw = "password"
    again:
    inp = InputBox("Enter Password")
    If inp = psw Then
        Application.Visible = True
        Me.Hide
    ElseIf StrPtr(inp) = 0 Then
        Exit Sub
    Else
        GoTo again
        
    End If
    End Sub
```

