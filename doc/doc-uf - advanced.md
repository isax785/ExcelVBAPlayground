# **Advanced Techniques**

- [**Advanced Techniques**](#advanced-techniques)
  - [**1 Passing parameters to a UserForm**](#1-passing-parameters-to-a-userform)
  - [**2 Returning values from a UserForm**](#2-returning-values-from-a-userform)
  - [**3 Dynamic control creation**](#3-dynamic-control-creation)
- [**Debugging UserForms**](#debugging-userforms)
- [**A Complete Mini‑Project Example**](#a-complete-miniproject-example)
    - [**UserForm Controls**](#userform-controls)
    - [**Initialization**](#initialization)
    - [**Submit**](#submit)



## **1 Passing parameters to a UserForm**
```vb
Public CustomerID As Long

Sub ShowFormWithID(id As Long)
    UserForm1.CustomerID = id
    UserForm1.Show
End Sub
```

---

## **2 Returning values from a UserForm**
```vb
' In UserForm
Public Result As String

Private Sub btnOK_Click()
    Result = txtName.Text
    Me.Hide
End Sub
```

```vb
' In module
Sub Test()
    UserForm1.Show
    MsgBox UserForm1.Result
End Sub
```

---

## **3 Dynamic control creation**
```vb
Dim txt As MSForms.TextBox

Set txt = Me.Controls.Add("Forms.TextBox.1", "txtDynamic", True)
txt.Top = 50
txt.Left = 20
txt.Width = 100
```

---

# **Debugging UserForms**

- ✔ Use `Debug.Print` generously  
- ✔ Test each control individually  
- ✔ Validate all user inputs  
- ✔ Use error handlers  

