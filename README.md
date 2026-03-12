# Readme

- [Readme](#readme)
  - [MOC](#moc)
  - [Contents](#contents)
  - [Utilities](#utilities)

---

A playground focused on the development of Excel applications for any type of engineering task.

The main goal is to speed-up and reduce the effort of implementations, while strenghtening the development skills.

## MOC

- [Documentation](./doc/doc-00_MOC.md): extensive dissertation of a specific topic.
- [Examples](./ex/ex-00_MOC.md): implementations (i.e. scripts) applied to real cases, ready to use or for demonstration purposes.
- [Playground](./plg/plg-00_MOC.md): Excel direct implementations for practicing and stranghening the development skills.
- [Sources](./src/src-00_MOC.md): sources and utilities.
- [Toolboxes](./tbx/tbx-00_MOC.md): quick reference (cheatsheet style) with reference tables and compact snippets.

## Contents

**Fields** this repo is focused on the implementation of:

- `Excel`: formula, array formula, and other functionalities;
- `Excel+VBA`: automation by VBA scripting;
- `UserForm`: GUIs implementation integrating both spreadsheet and VBA functionalities.

**File Naming**

Files are named by well-defined prefixes to ease the search of the requested file and the identification of the field it applies to:

| Field       | Identifier |
| ---         | ---        |
| `Excel`     | `xl`       |
| `Excel+VBA` | `vba`      |
| `UserForm`  | `uf`       |

Also filetypes have dedicated prefixes depending on the field of application:

| Filetype                                  | Identifier |
| ---                                       | ---        |
| Example                                   | `ex`       |
| Course                                    | `c`        |
| Document focusing on a specific topic     | `doc`      |
| Playground file                           | `plg`      |
| Toolbox                                   | `tbx`      |

Prefixes are built by combinating the identifiers reported in the tables above.

## Utilities

- [ToolBox Guidelines](./src/src-toolbox_guidelines.md)
- [Folder File Index Generator](./src/build_index.py)

**Table Template**

| Action                                 | Code                                  |
| ---                                    | ---                                   |
| some action here                       | `code`                                |