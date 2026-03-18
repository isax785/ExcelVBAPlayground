# Vba Error Handling Toolbox

| Action                                 | Code                                  |
| ---                                    | ---                                   |
| Switcher off error handling (until next _On Error_ statement)       | `On Error ()`             |
| Execution continues with the line following the error line          | `On Error Resume Next`    |
| Execution jumps to line starting with the specified label (+ colon) | `On Error GoTo [myLabel]` |
| Execution resumes with the statement that caused the error          | `Resume`                  |
| Execution resumes with the line following the error line            | `Resume Next`             |
| Execution resumes at the line starting with a specified label       | `Resume [myLabel]`        |
| Raise error                            | *`Err.Raise vbObjectError + [int], , [error-message]`* |
| **Error Properties** |   |
| Number  | `Err.Number`   |
| Description  | `Err.Description`   |
| Source  | `Err.Source`   |    

**General purpose error handler**

```vb
Sub AnySub()
    On Error GoTo ErrTrap
    ...
    Exit Sub ErrTrap:
    MsgBox "Number:" & Err.Number & vbCr & _
                     "Description: "  & Err.Description & vbCr * _
                     "Source: " & Err.Source, vbCritical, "some other message"
    
end Sub
```


---

[MOC](./tbx-00_MOC.md)