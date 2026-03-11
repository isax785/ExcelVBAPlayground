# Example: Actions

```vb
Sub Actions()
' Actions Macro
    ' write some values
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("B2:B3").Select
    ' autofill by dragging vertically
    Selection.AutoFill Destination:=Range("B2:B23"), Type:=xlFillDefault
    
    Range("B2:B23").Select
    ' autofill by dragging horizontally
    Selection.AutoFill Destination:=Range("B2:C23"), Type:=xlFillDefault
    
    ' selection with Ctrl + A
    Range("B2:C23").Select
    Range("B2:C23").Select
    
    ' copy-paste
    Selection.Copy
    Range("E2").Select
    ActiveSheet.Paste
    Range("B2").Select
    Application.CutCopyMode = False
    
    ' select with Ctrl+Arrow
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("E2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    
    Selection.Copy
    Range("H2").Select
    ActiveSheet.Paste
    Range("E4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Application.CutCopyMode = False
    
    ' cut-paste
    Selection.Cut
    Range("I4").Select
    ActiveSheet.Paste
End Sub
```