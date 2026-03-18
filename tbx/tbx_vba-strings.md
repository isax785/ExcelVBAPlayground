# VBA Toolbox - Strings


| **Manipulation**                |                                       |
| ---                                    | ---                                   |
| String length -> `int`                 | *`Len([str])`*                        |
| String concatenation (spaces are not automatically inserted) | *`"[string]" & "[string]" * [int]`* |
| Convert value to string                | *`CStr([val])`*                       |
| String to upper/lower case             | *`UCase([string])`* / *`LCase([string])`* |
| Character to integer                   | `CInt(...)`                           |
| Reverse a string -> `str`              | *`[str_rev] = StrReverse([str])`*     |
| Left characters                        | *`Left([str], [n_char])`*             |
| Right characters                       | *`Right([str], [n_char])`*            |
| Trim left, right, both-sides leading spaces | *`LTrim([string])` `RTrim([string])` `Trim([string])`* |
| Trim with high efficiency`*` | *`Trim$([string])`* |

`*` it works directly on strings instead of variants

---

[MOC](./tbx-00_MOC.md)
