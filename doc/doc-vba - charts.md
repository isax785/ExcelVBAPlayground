# Charts

Some useful resources:

- https://peltiertech.com/Excel/Charts/index.html
  - https://peltiertech.com/Excel/Charts/axes.html

## [Chart Axis](https://exceloffthegrid.com/chart-axis-min-mix/)

Set chart axis min and max based on a cell value.

- Updates automatically whenever data changes
- Does not require user interaction – i.e. no button clicking, but updates automatically when the worksheet recalculates
- Easily portable between different worksheets

```vb
Function setChartAxis(sheetName As String, chartName As String, MinOrMax As String, _
    ValueOrCategory As String, PrimaryOrSecondary As String, Value As Variant)

'Create variables
Dim cht As Chart
Dim valueAsText As String

'Set the chart to be controlled by the function
Set cht = Application.Caller.Parent.Parent.Sheets(sheetName) _
    .ChartObjects(chartName).Chart

'Set Value of Primary axis
If (ValueOrCategory = "Value" Or ValueOrCategory = "Y") _
    And PrimaryOrSecondary = "Primary" Then

    With cht.Axes(xlValue, xlPrimary)
        If IsNumeric(Value) = True Then
            If MinOrMax = "Max" Then .MaximumScale = Value
            If MinOrMax = "Min" Then .MinimumScale = Value
        Else
            If MinOrMax = "Max" Then .MaximumScaleIsAuto = True
            If MinOrMax = "Min" Then .MinimumScaleIsAuto = True
        End If
    End With
End If

'Set Category of Primary axis
If (ValueOrCategory = "Category" Or ValueOrCategory = "X") _
    And PrimaryOrSecondary = "Primary" Then

    With cht.Axes(xlCategory, xlPrimary)
        If IsNumeric(Value) = True Then
            If MinOrMax = "Max" Then .MaximumScale = Value
            If MinOrMax = "Min" Then .MinimumScale = Value
        Else
            If MinOrMax = "Max" Then .MaximumScaleIsAuto = True
            If MinOrMax = "Min" Then .MinimumScaleIsAuto = True
        End If
    End With
End If

'Set value of secondary axis
If (ValueOrCategory = "Value" Or ValueOrCategory = "Y") _
    And PrimaryOrSecondary = "Secondary" Then

    With cht.Axes(xlValue, xlSecondary)
        If IsNumeric(Value) = True Then
            If MinOrMax = "Max" Then .MaximumScale = Value
            If MinOrMax = "Min" Then .MinimumScale = Value
        Else
            If MinOrMax = "Max" Then .MaximumScaleIsAuto = True
            If MinOrMax = "Min" Then .MinimumScaleIsAuto = True
        End If
    End With
End If

'Set category of secondary axis
If (ValueOrCategory = "Category" Or ValueOrCategory = "X") _
    And PrimaryOrSecondary = "Secondary" Then
    With cht.Axes(xlCategory, xlSecondary)
        If IsNumeric(Value) = True Then
            If MinOrMax = "Max" Then .MaximumScale = Value
            If MinOrMax = "Min" Then .MinimumScale = Value
        Else
            If MinOrMax = "Max" Then .MaximumScaleIsAuto = True
            If MinOrMax = "Min" Then .MinimumScaleIsAuto = True
        End If
    End With
End If

'If is text always display "Auto"
If IsNumeric(Value) Then valueAsText = Value Else valueAsText = "Auto"

'Output a text string to indicate the value
setChartAxis = ValueOrCategory & " " & PrimaryOrSecondary & " " _
    & MinOrMax & ": " & valueAsText

End Function
```

Then on the spreadsheet:	![image-20201027165254402](img_md/image-20201027165254402.png)

**<u>Example</u>**

![image-20201127094538738](img_md/image-20201127094538738.png)

## Chart Range

Change chart formula:

```vb
Sub ChangeSeriesFormula()
    ''' Just do active chart
    If ActiveChart Is Nothing Then
        '' There is no active chart
        MsgBox "Please select a chart and try again.", vbExclamation, _
            "No Chart Selected"
        Exit Sub
    End If

    Dim OldString As String, NewString As String, strTemp As String
    Dim mySrs As Series

    OldString = InputBox("Enter the string to be replaced:", "Enter old string")

    If Len(OldString) > 1 Then
        NewString= InputBox("Enter the string to replace " & """" _
            & OldString & """:", "Enter new string")
        '' Loop through all series
        For Each mySrs In ActiveChart.SeriesCollection
            strTemp = WorksheetFunction.Substitute(mySrs.Formula, _
                OldString, NewString)
            mySrs.Formula = strTemp
        Next
    Else
        MsgBox "Nothing to be replaced.", vbInformation, "Nothing Entered"
    End If
End Sub
```

Assuming that you want to expand the range (by adding one extra column) to add one more observation for each series in you diagram (and not to add a new series), you could use this code:

```vb
Sub ChangeChartRange()
    Dim i As Integer, r As Integer, n As Integer, p1 As Integer, p2 As Integer, p3 As Integer
    Dim rng As Range
    Dim ax As Range

    'Cycles through each series
    For n = 1 To ActiveChart.SeriesCollection.Count Step 1
        r = 0

        'Finds the current range of the series and the axis
        For i = 1 To Len(ActiveChart.SeriesCollection(n).Formula) Step 1
            If Mid(ActiveChart.SeriesCollection(n).Formula, i, 1) = "," Then
                r = r + 1
                If r = 1 Then p1 = i + 1
                If r = 2 Then p2 = i
                If r = 3 Then p3 = i
            End If
        Next i


        'Defines new range
        Set rng = Range(Mid(ActiveChart.SeriesCollection(n).Formula, p2 + 1, p3 - p2 - 1))
        Set rng = Range(rng, rng.Offset(0, 1))

        'Sets new range for each series
        ActiveChart.SeriesCollection(n).Values = rng

        'Updates axis
        Set ax = Range(Mid(ActiveChart.SeriesCollection(n).Formula, p1, p2 - p1))
        Set ax = Range(ax, ax.Offset(0, 1))
        ActiveChart.SeriesCollection(n).XValues = ax

    Next n
End Sub
```

### Dynamic Range Cahnge without VBA

You don't need VBA to make the chart dynamic. Just create dynamic named ranges that grow and shrink with the data. Your VBA for the chart can refer to that named range, without adding burden to the code. But you may not even need VBA at all. The chart defined with dynamic ranges will update instantaneously. No code required.

Range name and formula for chart labels:

```
chtLabels =Sheet1!A2:Index(Sheet1!$A:$A,counta(Sheet1!$A:$A))
```

Range name and formula for column B

```vb
chtB =offset(chtlabels,0,1)
```

Range and formula for column D

```vb
chtD =offest(chtlabels,0,3)
```

Edit the data source and instead of the fixed cell ranges, enter the named ranges in the format

```
=Sheet1!*RangeName*
```

You need to supply the respective range name to the series values and the chart's category axis.

Note: when you supply dynamic range names as the source for a chart, you MUST include either the file name or the sheet name with the range name reference. When you close and re-open the dialog, you will find that Excel automatically converts your entry to the format `[Filename]RangeName`

Note 2: there are many different formula options to create dynamic range names. In this case, we're using and index of column A and determine the last populated cell by counting the cells. This only works if all cells in column A have text. If your data has gaps in column A (which I don't think you do), different formulas can be applied to determine the range.
