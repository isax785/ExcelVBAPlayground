# Vba Error Handling Toolbox

| Action                                 | Code                                  |
| ---                                    | ---                                   |
| Switch *off error* handling (till next `On Error` statement) | `On Error 0`   |
| Execution continues with the line following the error line | `On Error Resume Next`   |
| Execution jumps to line starting with the specified label (+ colon) | `On Error GoTo myLabel`   |
| Execution resumes with the statement that caused the error     | `Resume`   |
| Execution resumes with the line following the error            | `Resume Next`   |
| Execution resumes at the line startingg with a specified label | `Resume myLabel`   |
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