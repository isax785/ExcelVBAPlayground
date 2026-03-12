# VBA Toolbox - Charts

- [VBA Toolbox - Charts](#vba-toolbox---charts)
  - [Copy-Paste Chart](#copy-paste-chart)
  - [Set Chart](#set-chart)

---

| **Charts**                             |                                       |
| ---                                    | ---                                   |
| Declare and set new chart | `Dim oChart as Chart : Set oChart = Charts.Add`    |
| Count chart objects within a worksheet | *`[worksheet].ChartObjects.Count`* |


## Copy-Paste Chart

Copy a chart from a sheet and paste it as a picture into another sheet:

```vb
Set oChart = Charts.Add
oChart.SetSourceData oRange
oChart.ChartType = xlXYScatterLinesNoMarkers
Sheets(1).Select
ActiveChart.ChartArea.Copy
Sheets(2).Seelct
ActiveSheet.PasteSpecial Format:="Picture (JPEG)"
Selection.ShapeRange.ScaleWidth 0.8, msoFalse
Selection.ShapeRange.ScaleHeight 0.8, msoFalse
Selection.ShapeRange.IncrementLeft 100
Selection.ShapeRange.IncremetnTop 100
```

## Set Chart

```vb
Dim oWs as Worksheet : Set oWs = ActiveWorksheet

oWs.Shapes.AddChart2(240, x1XYScatterLines).Select
ActiveChart.SetSourceData oWs.Range(...)
ActiveChart.HasTitle = False
oWs.ChartObjects(1).Top = Range(...).Top
oWs.ChartObjects(1).Left = Range(...).Left
oWs.ChartObjects(1).Width = 300
oWs.ChartObjects(1).Height = 150
```

or into a `With` statement:

```vb
With ActiveChart
    .[property] = ...
```


---

[MOC](./tbx%20-%2000%20MOC.md)