# **A Complete Mini‑Project Example**

Here’s a compact example of a **data entry form**:

### **UserForm Controls**
- `txtName`
- `cmbCountry`
- `txtAge`
- `btnSubmit`
- `btnCancel`

### **Initialization**
```vb
Private Sub UserForm_Initialize()
    cmbCountry.List = Array("Italy", "France", "Germany", "Spain")
End Sub
```

### **Submit**
```vb
Private Sub btnSubmit_Click()

    If txtName.Text = "" Or txtAge.Text = "" Then
        MsgBox "All fields required.", vbExclamation
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = Sheets("People")

    Dim r As Long
    r = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1

    ws.Cells(r, 1).Value = txtName.Text
    ws.Cells(r, 2).Value = cmbCountry.Text
    ws.Cells(r, 3).Value = txtAge.Text

    MsgBox "Saved.", vbInformation
    Unload Me
End Sub
```