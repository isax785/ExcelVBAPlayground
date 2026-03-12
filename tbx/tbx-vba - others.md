# VBA TooBox - Others



| **Simulation Methods**                |                                       |
| ---                                    | ---                                   |
| Goal seek            | *`Range(...).GoalSeek [goal-value], [cell-to-change]`*  |
| Solver OK function by displaying the dialog | *`SolverOkDialog [SetCell], [MaxMinVal], [ValueOf], [ByChange], [Engine], [EngineDesc]`*  |
|                 | `SolverOkDialog "H6", 2, 0, "H1:H2", 1, "GRG Nonlinear"`   |


# Snippets



## Write Text File

```vb
    Dim fileSystemObject As Object
    Dim textStream As Object
    ...
    Set fileSystemObject = CreateObject("Scripting.FileSystemObject")
    Set textStream = fileSystemObject.CreateTextFile([filename], True, False)
    textStream.Write [string]
    ...
    textStream.Close
    ' Clean up
    Set fileSystemObject = Nothing
    Set textStream = Nothing
```



## Date

```vb
Dim dDate as DAte
dDate = DateSerial([year], [month], [day])
weekday = Weekday(dDate)
year = Year(dDate)
```
