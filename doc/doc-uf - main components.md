# **Main UserForm Components & Implementation Examples**

Below are the most common controls with **best practices** and **code snippets**.

- [**Main UserForm Components \& Implementation Examples**](#main-userform-components--implementation-examples)
- [**1 TextBox**](#1-textbox)
- [**2 ComboBox**](#2-combobox)
- [**3 ListBox**](#3-listbox)
- [**4 CheckBox**](#4-checkbox)
- [**5 OptionButton**](#5-optionbutton)
- [**6 CommandButton**](#6-commandbutton)
- [**7 MultiPage**](#7-multipage)
- [**8 Frame**](#8-frame)
- [**9 Labels**](#9-labels)

---

# **1 TextBox**
**Use cases**
- Names
- Numbers
- Dates
- Search fields

**Example: Numeric‑only TextBox**
```vb
Private Sub txtQuantity_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0
    End If
End Sub
```

**Example: Trim input on exit**
```vb
Private Sub txtName_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    txtName.Text = Trim(txtName.Text)
End Sub
```

---

# **2 ComboBox**

**Best practice**: Always load items dynamically.

**Example: Load from a sheet**
```vb
Private Sub UserForm_Initialize()
    Dim rng As Range
    Set rng = Sheets("Lists").Range("A2:A50").SpecialCells(xlCellTypeConstants)

    Dim cell As Range
    For Each cell In rng
        cmbCountry.AddItem cell.Value
    Next cell
End Sub
```

**Example: Load from an array**
```vb
Private Sub UserForm_Initialize()
    Dim arr
    arr = Array("Low", "Medium", "High")

    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        cmbPriority.AddItem arr(i)
    Next i
End Sub
```

---

# **3 ListBox**

**Use cases**
- Displaying lists
- Multi‑selection
- Search results

**Example: Populate ListBox from a table**
```vb
Private Sub LoadProducts()
    Dim tbl As ListObject
    Set tbl = Sheets("Data").ListObjects("Products")

    lstProducts.Clear
    lstProducts.ColumnCount = tbl.ListColumns.Count
    lstProducts.List = tbl.DataBodyRange.Value
End Sub
```

**Example: Get selected items**
```vb
Dim i As Long
For i = 0 To lstProducts.ListCount - 1
    If lstProducts.Selected(i) Then
        Debug.Print lstProducts.List(i, 0)
    End If
Next i
```

---

# **4 CheckBox**
**Example: Enable/Disable controls**
```vb
Private Sub chkAdvanced_Click()
    fraAdvanced.Enabled = chkAdvanced.Value
End Sub
```

---

# **5 OptionButton**
Use inside a **Frame** to group them.

**Example: Get selected option**
```vb
Dim gender As String

If optMale.Value Then gender = "Male"
If optFemale.Value Then gender = "Female"
```

---

# **6 CommandButton**
**Example: Submit button with validation**
```vb
Private Sub btnSubmit_Click()

    If Trim(txtName.Text) = "" Then
        MsgBox "Name is required.", vbExclamation
        txtName.SetFocus
        Exit Sub
    End If

    ' Process data
    Call SaveRecord

    MsgBox "Record saved.", vbInformation
    Unload Me
End Sub
```

---

# **7 MultiPage**

Great for wizards or complex forms.

**Example: Next/Back navigation**
```vb
Private Sub btnNext_Click()
    mpgMain.Value = mpgMain.Value + 1
End Sub

Private Sub btnBack_Click()
    mpgMain.Value = mpgMain.Value - 1
End Sub
```

---

# **8 Frame**

Use to group related controls visually and logically.

---

# **9 Labels**
Use for:
- Field names
- Status messages
- Instructions

**Example: Dynamic status label**
```vb
lblStatus.Caption = "Loading data..."
```
