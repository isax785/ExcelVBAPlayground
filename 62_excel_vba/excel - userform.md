# USER FORM

## CREATE USER FORM

In Macro Excel, click on "Insert" --> "UserForm"

Insert tools and controls on the interface. Double click on the controls in order to program them.

### INITIALIZATION

Actions to do when the UserForm is initialized.

```
    Private Sub UserForm_Initialize()
    	...
    End Sub
```
### CREATE A LIST FROM A RANGE IN A WORKSHEET

If range in column:

```
    Me.ComboBox1.List = Range("A1:A6").Value  
```
If range in row:

```
    Me.ComboBox1.List = Application.Transpose(Range("A1:G1").Value)  
```

Other methods:

```
    Me.ComboBox1.List = Range(Selection, Selection.End(xlDown)).Value  
```

If merged cells in the range:

```
    Dim aCell As Range
    ...
    
    For Each aCell In Selection
        If aCell.Value <> "" Then
            Me.ComboBox1.AddItem aCell.Value
        End If
    Next
```
### COMBO BOX   

```
    Private Sub ComboBox1_Change()
    ...
    End Sub
```
### COMMAND BUTTON  

```
    Private Sub CommandButton1_Click()
    ...
    End Sub
```

## HOW TO CALL A USER FORM

### CALL FROM A COMMAND IN THE EXCEL FILE

In a worksheet, "Develop" --> "Insert" --> ActiveX Controls --> Command Button

Double click on the Button.

```
    Private Sub StartConversion_Click()
    	UserForm1.Show
    End Sub
```
## OPEN USER FORM WHEN OPEN THE EXCEL FILE 

Also, to hide the excel application.

```
    Private Sub Workbook_open()
    	Application.Visible = False
    	UserForm1.Show
    End Sub
```
## WHAT TO DO WHEN CLOSE A USER FORM

Make the excel application visible, close the workbook, ...

Example: close the workbook if the excel application is not visible (general case of generic user, not administrator).

```
    Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
        If Application.Visible = False Then
            Workbooks.Close
        End If
    End Sub
```

## MAKE THE EXCEL APPLICATION VISIBLE BY A COMMAND BUTTON AND PSW

```
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

