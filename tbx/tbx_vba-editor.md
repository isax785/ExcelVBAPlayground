# VBA Toolbox - Editor

| What               | Shortcut  | Description                            |
| ---                | ---       | ---                                    |
| Editor             | `Alt+F11` |  open macro editor.                    | 
| Project Explorer   | `Ctrl+R`  |  structure of workbooks/modules/forms. |
| Properties Window  | `F4`      |  object properties.                    |
| Code Window        |           |  editor, procedures list.              |
| Immediate Window   | `Ctrl+G`  |  quick evaluation `?ActiveSheet.Name`, `Debug.Print`, one‑off commands. |
| Locals Window      |           |  inspect variables in current scope.          |
| Watch Window       |           |  track expressions; break when value changes. |
| Object Browser     | `F2`      |  discover classes, methods, constants.        |


**Immediate Window** 
- `?` operator to get a variable/property value
- `=` operator to set/assign a variable/property value

| Action                           | Code                      |
| ---                              | ---                       |
| Compare two values               | `? 5>4`                   | 
| Find the name of a sheet         | `?Sheet1.Name`            |
| Find cell calue                  | `?Range("A1").Value`      |
| Find active cell calue           | `?ActiveCell.Value`       |
| Count number of sheets in the active workbook | `?Sheets.Count` |
| Assign name to a sheet           | `sheet1.Name = "SomeName"`   |
| Assign a value to a cell         | `Range("A1").Value = 12`     |
| Debug print (within the script)  | `Debug.Print "content"`      |
| Run function defined in a Module | *`?[function_name]([arguments])`* |
| Run macro                        | *`[macro_name]([arguments])`* |



---

[MOC](./tbx-00_MOC.md)