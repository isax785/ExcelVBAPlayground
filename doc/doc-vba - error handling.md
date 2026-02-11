# Error Handling

| Command                 | Action                                                       |
| ----------------------- | ------------------------------------------------------------ |
| On Error ()             | Switcher off error handling (until next _On Error_ statement) |
| On Error Resume Next    | Execution continues with the line following the error line   |
| On Error GoTo *myLabel* | Execution jumps to line starting with the specified label (+ colon) |
| Resume                  | Execution resumes with the statement that caused the error   |
| Resume Next             | Execution resumes with the line following the error line     |
| Resume _myLabel_        | Execution resumes at the line starting with a specified label |

Example of a general error handler:

```vb
Sub AnySub()
	On Error GoTo ErrTrap
	....
	Exit Sub
	ErrTrap:
        MsgBox "Error Message"
End Sub
```