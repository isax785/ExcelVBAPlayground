# Data Validation Tips and Quirks

Tips for solving problems with drop down lists, and workarounds for data validation limitations, such as small font, and narrow lists:

- [Refer to a Source List on a Different Worksheet](https://www.contextures.com/xlDataVal08.html#Refer) 
- [Use Dynamic Lists](https://www.contextures.com/xlDataVal08.html#Dynamic)
- [Item Limit in Drop Down List](https://www.contextures.com/xlDataVal08.html#itemlimit)
- [Drop Down List Opens With Blank Selected](https://www.contextures.com/xlDataVal08.html#blankcell)
- [Video: Dynamic Range](https://www.contextures.com/xlDataVal08.html#videodynamic)
[Data Validation Font Size and List Length](https://www.contextures.com/xlDataVal08.html#Font)
- [Data Validation List With Symbols](https://www.contextures.com/xlDataVal08.html#symbols)
- [Scroll Through a Drop Down List](https://www.contextures.com/xlDataVal08.html#scroll)
- [Data Validation Dropdowns and Change Events](https://www.contextures.com/xlDataVal08.html#Change)
- [Missing Arrows](https://www.contextures.com/xlDataVal08.html#ArrowsNotVisible) 
- [Valid Entries Not Allowed](https://www.contextures.com/xlDataVal08.html#valid)
- [Invalid Entries are Allowed](https://www.contextures.com/xlDataVal08.html#Invalid) 
  - [Video: Ignore Blank in Data Validation](https://www.contextures.com/xlDataVal08.html#videoignore)
- [Data Validation on a Protected Sheet](https://www.contextures.com/xlDataVal08.html#Protect) 
- [Data Validation Dropdowns are Too Wide](https://www.contextures.com/xlDataValWidth.html)  
- [Make the Dropdown List Temporarily Wider](https://www.contextures.com/xlDataVal08.html#Wider)
- [Make the Dropdown List Appear Larger](https://www.contextures.com/xlDataVal08.html#Larger)

## Introduction

To create a drop down list in a worksheet cell, use Excel's data validation feature. There are basic instructions on [the Getting Started page](https://www.contextures.com/xlDataVal01.html), and many other techniques, such as [Dependent Drop Down Lists](https://www.contextures.com/xlDataVal02.html), and [showing a popup Combo Box](https://www.contextures.com/xlDataVal10.html) when a data validation cell is clicked.

There are many tips in the following sections, for working efficiently with data validation, and troubleshooting tips for when things go wrong.

![Dependent drop-down](https://www.contextures.com/images/datavalidation/datavaldependfruit.png)

## Refer to Source List on Different Sheet

When you try to create an Excel data validation dropdown list, and refer to a source list on a different worksheet, you might see an error message: "*You may not use references to other worksheets or workbooks for Data Validation criteria.*"

![Refer to a Source List on a Different Worksheet](https://www.contextures.com/images/DV92.gif)

To avoid this problem, name the list on the other worksheet, then refer to the named range, as described here:  [Excel Data Validation](https://www.contextures.com/xlDataVal01.html) 

If the list is in a different workbook, you can use the technique described here: [Use a List from Another Workbook](https://www.contextures.com/xlDataVal05.html)

## Use Dynamic Lists

Some lists change frequently, with items being added or removed. If the list is the source for a Data Validation dropdown, use a dynamic formula to name the range, and the dropdown list will be automatically updated.

For instructions, view this page:   [Create a Dynamic Range](https://www.contextures.com/xlNames01.html#Dynamic)[![go to top](https://www.contextures.com/images/scrollup.gif)](https://www.contextures.com/xlDataVal08.html#Top)

## Item Limit in Drop Down List

There are limits to the number of items that will show in a data validation drop down list:

- The list can show up to show **32,767 items** from a list on the worksheet.
- If you type the items into the data validation dialog box (a delimited list), the limit is **256 characters**, including the separators.

If you need more items than that, you could create a dependent drop down list, broken down by category. There is a sample file here: [Dependent Drop Down from Sorted List](https://www.contextures.com/xlDataVal13.html)

## Drop Down List Opens With Blank Selected

When you click the arrow to open a drop down list, the selection might go to a blank at the bottom of the list, instead of the first item in the list.

To download the sample file, click here: [Remove Blanks With Dynamic Range Sample File](https://www.contextures.com/ExcelTemplates/DataValBlankSelect.zip)

Why does this happen, and how can you prevent it? Also, if there are blanks in the source list, [invalid entries might be allowed in the cells](https://www.contextures.com/xlDataVal08.html#blank).

![drop down list blank](https://www.contextures.com/images/datavalidation/dropdownlistblank.png)

In the example shown above, the drop down list is based on a range named Products. The person who set up the list left a few blank cells at the end, where new items could be added.

![drop down list blank 02](https://www.contextures.com/images/datavalidation/dropdownlistblank02.png)

If there's a blank cell in the source list, and the cell with the data validation list is blank, the list will open with the blank entry selected.

### Prevent the Problem

To prevent this, either enter a default value in the data validation cell, or remove the blanks from the source list.

#### Create a Default Item at Top of List

Or, make " --Select--" the top item in the Product list, and set up the worksheet with " --Select--" entered in each product cell, as the default entry.

**NOTE**: Type a space or an apostrophe at the start of "--Select--" so Excel will not show you an error message.

![drop down list blank 03](https://www.contextures.com/images/datavalidation/dropdownlistblank03.png)

#### Remove Blanks With Dynamic Named Range

So, in this example, you could change the Products list to a [dynamic range](https://www.contextures.com/xlNames01.html#Dynamic), which will adjust automatically when items are added or removed.

The OFFSET formula used in this example is:

**=OFFSET(Prices!$B$2,0,0,COUNTA(Prices!$B:$B)-1,1)**

![drop down list blank 03](https://www.contextures.com/images/datavalidation/dropdownlistblank07.png)

## Watch the Dynamic Range Video

To see the steps for setting up a dynamic named range, please watch this short video tutorial.



## Data Validation Font Size and List Length

The font type and font size in a data validation list can't be changed, nor can its default list length, which has a maximum of eight rows.

If you reduce the zoom setting on a worksheet, it can be almost impossible to read the items in the dropdown list, as in the example below.

![almost impossible to read the items](https://www.contextures.com/images/DV66.gif)

One workaround is to use programming, and a combo box from the Control Toolbox, to overlay the cell with data validation. If the user double-clicks on a data validation cell, the combobox appears, and they can choose from it. There are [instructions here](https://www.contextures.com/xlDataVal10.html).[![go to top](https://www.contextures.com/images/scrollup.gif)](https://www.contextures.com/xlDataVal08.html#Top)

![combo box from the Control Toolbox](https://www.contextures.com/images/DV67.gif)

## Data Validation List With Symbols

Unfortunately, you can't change the font type or font size in a data validation list, as mentioned in the previous section. The drop down list always shows **Tahoma font**, even if the source list is in a different font, such as Wingdings or Symbol.

However, you can use symbol characters from the Tahoma font, such as arrows, circles, and squares.

![drop down list with symbols](https://www.contextures.com/images/datavalidation/symbollist01.png)

This video shows the steps to show symbols, and the written instructions are below the video.



To create a list of symbols:

1. On the worksheet, select a cell where you want to start the list of symbols
2. Press the Alt key, and on the number keypad, type a number for the symbol that you want to insert. A few examples are shown in the list below, and you can experiment to find other symbols.
   Note: To see all the codes, go to the [Alt Codes List](https://en.wikipedia.org/wiki/Template:Alt#Codes) in Wikipedia.
3. ![drop down list with symbols](https://www.contextures.com/images/datavalidation/symbollist02.png)

4. Press Enter, and enter other symbols in the cells below. In the list shown above, the Alt key was used with numbers 30, 29 and 31, to create a list with up and down arrows, and a two-headed arrow.

To create a drop down list with the symbols:

1. Select the cell where you want the drop down list
2. On the Ribbon's Data tab, click Data Validation
3. From the Allow drop down, select List
4. Click in the Source box, and on the worksheet, select the cells with the list of symbols, then click OK

To see the example, you can download the sample file: [Data Validation List With Symbols](https://www.contextures.com/datavalidationsamples/datavalsymbols.zip)

## Scroll Through a Drop Down List

After you create a drop down list, click on that cell, to see its drop down arrow. The arrow will only show when the cell is active.

To [open the drop down list](https://www.contextures.com/xlDataVal08.html#showlist), and to [scroll through the items](https://www.contextures.com/xlDataVal08.html#scrolllist) in the drop down list, you can use the mouse or the keyboard

NOTE: The list will only show 8 items at a time.

![drop down list with symbols](https://www.contextures.com/images/datavalidation/datavalidationscroll01.png)

#### Show the Drop Down List

- Mouse: Click the cell's arrow
- Keyboard: Press Alt + Down Arrow

#### Scroll Through the List Items

- Mouse:

- - Press the arrows at the top or bottom of the scroll bar, for continuous scrolling
  - Click the arrows at the top or bottom of the scroll bar, to scroll one item at a time
  - Drag the scroll box up or down
  - Click above or below the scroll box, to move up or down one page
  - Press above or below the scroll box, for continuous page scrolling

- Keyboard: Press Alt + Down Arrow

- - Press the Up or Down Arrows keys, for continuous scrolling
  - Tap the Up or Down Arrows keys, to scroll one item at a time
  - Tap the Home or End key, to go to the top or bottom of the list
  - Tap the Page Up or Page Down key, to move up or down one page
  - Press the Page Up or Page Down key, for continuous page scrolling

## Data Validation Dropdowns and Change Events

In Excel 2000 and later versions, selecting an item from a Data Validation dropdown list will trigger a Change event. This means that code can automatically run after a user selects an item from the list.

To see an example, go to the [Sample Worksheets](https://www.contextures.com/excelfiles.html) page, and under the **Filters** heading, find **Product List by Category**, and download the ***ProductsList.zip*** file.

In Excel 97, selecting an item from a Data Validation dropdown list **does not** trigger a Change event, unless the list items have been typed in the Data Validation dialog box. In this version, you can add a button to the worksheet, and run the code by clicking the button. To see an example, go to the [Sample Worksheets](https://www.contextures.com/excelfiles.html) page, and under the **Filters** heading, find **Product List by Category**, and download the **ProductsList97.zip** file.

Another option in Excel 97 is to use the Calculate event to run the code. To do this, refer to the cell with data validation in a formula on the worksheet, e.g. **=MATCH(C3,CategoryList,0)**. Then, add the filter code to the worksheet's Calculate event. To see an example, go to the [Sample Worksheets](https://www.contextures.com/excelfiles.html)page, and under the **Filters** heading, find **Product List by Category**, and download the **ProductsList97Calc.zip** file.[![go to top](https://www.contextures.com/images/scrollup.gif)](https://www.contextures.com/xlDataVal08.html#Top)

## Missing Arrows

Occasionally, data validation dropdown arrows are not visible on the worksheet, in cells where you know that data validation lists have been created.

This video shows the most common reasons for missing arrows. Written instructions for fixing the problems are below the video.



Here are a few causes of missing arrow for data validation. Click a link to see the details:

[Active Cell Only](https://www.contextures.com/xlDataVal08.html#activecell)

[Hidden Objects](https://www.contextures.com/xlDataVal08.html#hidden)

[Dropdown Option](https://www.contextures.com/xlDataVal08.html#option)

[Freeze Panes](https://www.contextures.com/xlDataVal08.html#Freeze)

[Corruption](https://www.contextures.com/xlDataVal08.html#corrupt)

[Deleted by Macro](https://www.contextures.com/xlDataVal08.html#delmacro)

### Active Cell Only

Only the active cell on a worksheet will display a data validation dropdown arrow. To mark cells that contain data validation lists, you can colour the cells, or add a comment.

If you require visible arrows for all cells that contain lists, you can use combo boxes instead of data validation, and those arrows will be visible at all times. To create a combo box:

- Click the Developer tab on the ribbon, and click Insert
- Click the Combo Box in the Form Controls
- On the worksheet, drag to add a combo box in the size that you want.
- Right-click the combo box, and click Format Control
- In the Input Range box, enter the name or address of the list
- Click OK

![combo box control](https://www.contextures.com/images/datavalidation/datavalcomboarrow01.png)

### Hidden Objects

If objects are hidden on the worksheet, the data validation dropdown arrows will also be hidden.

To make objects visible, use the keyboard shortcut -- **Ctrl + 6**

Or, follow these steps, to change the Option settings:

- Click the File tab on the ribbon, and click Options
- Click the Advanced category
- Scroll down about halfway, to the section, Display Options for This Workbook .
- In the setting, "For Objects, show:", click All
- Click OK

![Display Options for This Workbook](https://www.contextures.com/images/datavalidation/optionsshowobjects01.png)

### Dropdown Option

In the Data Validation dialog box, you can turn off the option for a dropdown list. To turn it back on:

1. Select the cell that contains a data validation list
2. On the Ribbon, click the Data tab
3. Click the top of the Data Validation button, to open the dialog box
4. On the Settings tab, add a check mark to In-cell dropdown
5. Click OK

![add a check mark to In-cell dropdown](https://www.contextures.com/images/datavalidation/datavaldropdown01.png)

### Excel 2013 Windows 8

In you have a linked picture in an Excel 2013 workbook, on Window 8, the data validation arrow might not appear in the active cell, unless you are pressing the mouse button.

![no arrow Excel 2013](https://www.contextures.com/images/datavalidation/datavalidationarrow01.png)

As a workaround, follow these steps to make the arrow appear:

1. Select the cell with the data validation list
2. Click outside of the Excel window (e.g. click on the Desktop, or click in your browser window)
3. Click on the Excel window, and the arrow will appear, and you can select an item from the list.

![no arrow Excel 2013](https://www.contextures.com/images/datavalidation/datavalidationarrow02.png)

### Freeze Panes

The Freeze Panes setting can cause problems with drop down arrows, in [all versions](https://www.contextures.com/xlDataVal08.html#freezeall) of Excel. There were additional problems in [Excel 97 and earlier](https://www.contextures.com/xlDataVal08.html#freeze97).

In any version of Excel, if a drop down list is in a frozen pane of the Excel window, and the column to the right has been scrolled off screen, the drop down arrow will not be visible.

Thanks to John Constable for this tip.

![no arrow Excel 2013](https://www.contextures.com/images/datavalidation/missingarrows01.png)

In Excel 97, if a Data Validation dropdown list is in a frozen pane of the window, the dropdown arrow does not appear when the cell is selected. As a workaround, use Window|Split instead of Window|Freeze Panes

This problem has been corrected in later versions.

![drop down without frozen panes](https://www.contextures.com/images/DV61.gif)Without frozen panes

![drop down with frozen panes](https://www.contextures.com/images/DV62.gif)
With frozen panes

### Corruption

If none of the above solutions explains the missing dropdown arrows, the worksheet may be corrupted. Try copying the data to a new worksheet or workbook, and the dropdown arrows may reappear.

Or, try to repair the file as you open it:

1. On the Ribbon, click File, and then click Open
2. Click Computer, then click Browse
3. Select the file with the missing data validation arrows
4. At the bottom of the Open windown, click the arrow at the right of he Open button
5. Click Open and Repair
6. When prompted, click Repair. [![go to top](https://www.contextures.com/images/scrollup.gif)](https://www.contextures.com/xlDataVal08.html#Top)

![open and repair](https://www.contextures.com/images/datavalidation/openandrepair01.png)

### Deleted by Macro

If you run a macro that deletes all the shapes on a worksheet, it might also delete the drop down arrow for data validation. Thanks to Ed Howland who suggested adding this tip.

For example, the macro below deletes all the shapes on the active sheet.

- If the **data validation arrow is visible** when you run this macro, it will be deleted too, along with other shapes on the worksheet.

**Safe Macros**: To delete other shapes safely, without deleting the data validation arrows, see the [macros to delete objects](https://www.rondebruin.nl/win/s4/win002.htm) on Ron de Bruin's website.

```
Sub DeleteShapesALL()
'WARNING: Deletes data val arrow
'         if it is visible
Dim sh As Shape
Dim ws As Worksheet
Set ws = ActiveSheet
For Each sh In ws.Shapes
  sh.Delete
Next sh
End Sub
```

## Valid Entries Not Allowed

If you type a valid entry in a cell that has a drop down list, you still might see an error message, stating that "The value you entered is not valid."

For example, this list allows you to choose Yes or No.

![list with Yes or No](https://www.contextures.com/images/datavalidation/notvalid02.png)

However, if you type "no", it is not valid.

![case sensitive delimited list](https://www.contextures.com/images/datavalidation/notvalid01.png)

This error can occur if the list is [based on a delimited list](https://www.contextures.com/xlDataVal01.html#Delimited), that is typed into the Data Validation dialog box.

This method of Data Validation is **case sensitive**, so you can choose from the drop down list, or type an entry that exactly matches the upper and lower case letters in the delimited list.

If you type "No", the entry will be accepted, without an error message, because the first letter is upper case, and the second letter is lower case.

![items in delimited list are case sensitive](https://www.contextures.com/images/datavalidation/notvalid03.png)

## Invalid Entries Are Allowed

Although you have created data validation dropdown arrows on some cells, users [may be able to type invalid entries](https://contexturesblog.com/archives/2010/06/25/invalid-entries-allowed-in-data-validation/). The following are the most common reasons for this.

To download the sample file, click here: [Data Validation Invalid Entries Sample File](https://www.contextures.com/datavalidationsamples/datavalinvalid.zip)

### Blank Cells in Source List

If the source list is a **named range** that contains blank cells, users may be able to type any entry, without receiving an error message. Watch this short video, to see one possible solution to the problem, or read the instructions below the video.



In the screen shot below, the Manager column has a drop down list with 5 names.

![drop down list of names](https://www.contextures.com/images/datavalidation/datavalblanks01.png)

However, if a different name is typed in that column, there is no error alert. The name Bill is not in the list, but was allowed in the cell.

![invalid name allowed](https://www.contextures.com/images/datavalidation/datavalblanks02.png)

This occurs when a [named range](https://www.contextures.com/xlNames01.html) is used as the list source, and there is a blank cell anywhere in that named range. Shown below is the named range, MgrList, with a blank cell at the end.

**Note**: If the source list is a **range address**, e.g. $A$1:$A$10, and contains blank cells, invalid entries will be blocked, with *Ignore blank* on or off.

![blank cell in named range](https://www.contextures.com/images/datavalidation/datavalblanks03.png)

To turn prevent this:

1. Select the cell that contains a data validation list
2. Choose Data|Validation
3. On the Settings tab, remove the check mark from the *Ignore blank* box.
4. Click OK

![ignore blank off](https://www.contextures.com/images/datavalidation/datavalblanks04.png)

### Video: Ignore Blank in Data Validation

Blank cells can also cause problems for dependent drop down lists. Watch this short Excel tutorial video on the potential problems when Ignore Blank is turned off, and the Circle Invalid Data feature is used. [![go to top](https://www.contextures.com/images/scrollup.gif)](https://www.contextures.com/xlDataVal08.html#Top)

 

### Error Alert

If the Error Alert is turned off, users will be able to type any entry, without receiving an error message. To turn the alert on:

1. Select the cell that contains a data validation list
2. Choose Data|Validation
3. On the Error Alert tab, add a check mark to the *Show error alert after invalid data is entered*box.
4. Click OK

![error alert off](https://www.contextures.com/images/datavalidation/datavalerroralertoff.png)

## Data Validation on a Protected Sheet

In Excel 2000 and earlier versions, you can change the selection in a data validation dropdown, if the list is from a range on the worksheet. If the list is typed in the data validation dialog box, the selection can't be changed.

In Excel 2002 and later versions, neither type of dropdown list can be changed if the cell is locked and the sheet is protected.

This MSKB article has information on the previous behaviour:

XL97: Error When Using Validation Drop-Down List Box https://support.microsoft.com/default.aspx?id=157484

## Make the Dropdown List Temporarily Wider

The Data Validation dropdown is the width of the cell that it's in, to a minimum of about 3/4". You could use a SelectionChange event to temporarily widen the column when it's active, then make it narrower when you select a cell in another column.

![Make the Dropdown List Temporarily Wider](https://www.contextures.com/images/datavalidation/datavalwider01.png)

For example, with Data Validation cells in column A:

```
Private Sub Worksheet_SelectionChange(ByVal Target As Range) 
  If Target.Count > 1 Then Exit Sub
   If Target.Column = 1 Then
       Target.Columns.ColumnWidth = 20
   Else
       Columns(1).ColumnWidth = 5
   End If 
End Sub 
```

**To add this code to the worksheet:**

1. Right-click on the sheet tab, and choose View Code.
2. Copy the code, and paste it onto the code module.
3. Change the column reference from 4 to match your worksheet.[![go to top](https://www.contextures.com/images/scrollup.gif)](https://www.contextures.com/xlDataVal08.html#Top)

## Make the Dropdown List Appear Larger

In a Data Validation dropdown list, you can't change the font or font size. So, if your worksheet is zoomed down, to show more cells, it will be difficult to read the drop down list, in its 8-point size. To address the problem, you could

- use a combo box or listbox, or
- temporarily increase the zoom setting.

### Use Combo Box or ListBox

To make it easier to read, you could use a combo box or listbox, to show the entries. The font in those can be set to any size, and you can also set them to show more than the default 8 items at a time. See instructions for [adding a combo box](https://www.contextures.com/xlDataVal14.html), or [showing a listbox](https://www.contextures.com/excel-data-validation-listbox.html) (can be set for single selection or multiple selection).

### Macro to Temporarily Change Zoom Setting

To make the text appear larger, you can use an event procedure (three examples are shown below) to increase the zoom setting when the cell is selected. (Note: this can be a bit jumpy)

![Macro to Temporarily Change Zoom Setting](https://www.contextures.com/images/datavalidation/datavallarger01.png)

Or, you can use code to display a combobox, as described in the [previous section](https://www.contextures.com/xlDataVal08.html#Font).

### Zoom in when specific cell is selected

If cell A2 has a data validation list, the following code will change the zoom setting to 120% when that cell is selected.

```
Private Sub Worksheet_SelectionChange(ByVal Target As Range) 
  If Target.Address  = "$A$2" Then 
    ActiveWindow.Zoom = 120 
  Else 
    ActiveWindow.Zoom = 100 
  End If 
End Sub 
```

**To add this code to the worksheet:**

1. Right-click on the sheet tab, and choose View Code.
2. Copy the code, and paste it onto the code module.
3. Change the cell reference from $A$2 to match your worksheet.[![go to top](https://www.contextures.com/images/scrollup.gif)](https://www.contextures.com/xlDataVal08.html#Top)

![macro to change zoom](https://www.contextures.com/images/DV60.gif)

### Zoom in when specific cells are selected

If several cells have a data validation list, the following code will change the zoom setting to 120% when any of those cells are selected. In this example, cells A1, B3 and D9 have data validation.

```
Private Sub Worksheet_SelectionChange(ByVal Target As Range) 
  If Target.Cells.Count > 1 Then Exit Sub
  If Intersect(Target, Range("A1,B3,D9")) Is Nothing Then 
    ActiveWindow.Zoom = 100 
  Else 
    ActiveWindow.Zoom = 120 
  End If 
End Sub  
```

[![go to top](https://www.contextures.com/images/scrollup.gif)](https://www.contextures.com/xlDataVal08.html#Top)

The following code will change the zoom setting to 120% when any cell with a data validation list is selected.

```
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
  Dim lZoom As Long
  Dim lZoomDV As Long
  Dim lDVType As Long
  lZoom = 100
  lZoomDV = 120
  lDVType = 0

  Application.EnableEvents = False
  On Error Resume Next
  lDVType = Target.Validation.Type
  
    On Error GoTo errHandler
    If lDVType <> 3 Then
      With ActiveWindow
        If .Zoom <> lZoom Then
          .Zoom = lZoom
        End If
      End With
    Else
      With ActiveWindow
        If .Zoom <> lZoomDV Then
          .Zoom = lZoomDV
        End If
      End With
    End If

exitHandler:
  Application.EnableEvents = True
  Exit Sub
errHandler:
  GoTo exitHandler
End Sub 
```