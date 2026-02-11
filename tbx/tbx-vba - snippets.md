# VBA Toolbox - Snippets

- [VBA Toolbox - Snippets](#vba-toolbox---snippets)
  - [Copy-Paste](#copy-paste)
  - [Delete Row](#delete-row)

---

## Copy-Paste

Copy from a sheet and paste values/formats/formulas into another one:

```vb
Worksheets(1).Cells(i, 3).Copy
' values
Worksheets(2).Cells(a, 15).PasteSpecial Paste:=xlPasteValues
' format
Worksheets(2).Cells(a, 15).PasteSpecial Paste:=xlPasteFormats
' formula
Worksheets(2).Cells(a, 15).PasteSpecial Paste:=xlPasteFormulas
'Disable marching ants around copied range 
Application.CutCopyMode = False
```

## Delete Row

Delete the row of the selected cell:

```vb
Sub find_delete_row() 
    Dim r As Long 
    'The panel Find must be open 
    Cells.Find(What:="TImestamp", After:=ActiveCell, LookIn:=xlFormulas, _ 
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _ 
        MatchCase:=False, SearchFormat:=False).Activate 
         
    Debug.Print ActiveCell.Row 
    r = ActiveCell.Row 
    Rows(r).Select 
        Selection.Delete Shift:=xlUp   'Shift up the cells below
End Sub
```