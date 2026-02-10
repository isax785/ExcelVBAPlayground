# VBA Snippets

- [VBA Snippets](#vba-snippets)
  - [Copy-Paste](#copy-paste)

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