# VBA Toolbox - Dialogs

- [VBA Toolbox - Dialogs](#vba-toolbox---dialogs)
- [Snippets](#snippets)
  - [Dialog for Folder Selection](#dialog-for-folder-selection)

---


| **MessageBox**                         |                                       |
| ---                                    | ---                                   |
| Open messagebox                        | *`MsgBox "[message]", [button-set], [box-title]`* |
| Messagebox button set                  | `vbOkCancel`, `vbYesNoCancel`, `vbYesNo` |
| Buttons signals                        | `vbOK`, `vbCancel`, `vbYes`, `vbNo`   |
| Conditional | *`If MsgBox("[message]", [button-set]) = [signal] Then [action] `* |
| Get messagebox output | `Dim msg as Variant` `msg = MsgBox(...)`               |
| Oputput cases                          | `Yes`    -> `Case 6`                  |
|                                        | `No`     -> `Case 7`                  |
|                                        | `Cancel` -> `Case 2`                  |


| **InputBox**                           |                                       |
| ---                                    | ---                                   |
| Open input box                        | *`Dim v as [type] : v = InputBox("[message]", ,[default])`* |
| | *`Application.InputBox (Prompt, Title, Default, Left, Top, HelpFile, HelpContextID, Type)`* |

# Snippets

## Dialog for Folder Selection

```vb
    Dim folderPath As String
    ...
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder to Save CSV Files"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With
```


---

[MOC](./tbx-00_MOC.md)