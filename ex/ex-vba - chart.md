# Example: Chart

```vb
Dim oChart as Chart

With oChart
    .ChartType = xlXYScatterSmooth
    .SetSourceData oRange
    .HasTitle = True
    .HasLegend = True
    .Axes(xlCategory).HasMajorGridlines = True
    .Axes(xlCategory).HasMinorGridlines = True
    pMin1 = WorksheetFunction.Min(Columns(1))
    pMin2 = WorksheetFunction.Min(Columns(2))
    .Axes(xlValue).MinimumScale = IIf(pMin1 < pMin2, Int(pMin1), Int(pMin2))
    .Axes(xlValue).HasMinorGridlines = True
    .ChartTitle.Caption = "Title"
    .SeriesCollection(1).Name = Cells(1, 1)
    .SeriesCollection(2).Name = Cells(1, 2)
End With

' Series and formatting
With oChart.SeriesCollection.NewSeries
    .XValues == Range("E5:E7")
    .Values == Range("F5:F7")
    .Name - "Insert"
    .ChartType = xlXYScatterLines
    .HasDataLabels = True
    .DataLabels.Select
    Selection.ShowCategoryName = True
    Selection.ShowValue = True
End With
```

