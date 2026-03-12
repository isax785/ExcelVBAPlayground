# **Comprehensive Guide to Filesystem Interaction in VBA**

- [**Comprehensive Guide to Filesystem Interaction in VBA**](#comprehensive-guide-to-filesystem-interaction-in-vba)
- [1. **Fundamentals of Filesystem Interaction**](#1-fundamentals-of-filesystem-interaction)
    - [**A. Native VBA built‑ins**](#a-native-vba-builtins)
    - [**B. FileSystemObject (FSO)**](#b-filesystemobject-fso)
- [2. **Path Handling in VBA**](#2-path-handling-in-vba)
  - [**2.1 Get current working directory**](#21-get-current-working-directory)
  - [**2.2 Change working directory**](#22-change-working-directory)
  - [**2.3 Extract components from a path**](#23-extract-components-from-a-path)
    - [**Get file name from full path**](#get-file-name-from-full-path)
    - [**Get folder portion of path**](#get-folder-portion-of-path)
    - [**Combine folder + filename safely**](#combine-folder--filename-safely)
- [3. **Folder Operations (Native VBA)**](#3-folder-operations-native-vba)
  - [**3.1 Create a folder**](#31-create-a-folder)
  - [**3.2 Remove a folder (only if empty)**](#32-remove-a-folder-only-if-empty)
  - [**3.3 Check if a folder exists**](#33-check-if-a-folder-exists)
- [4. **Folder Operations Using FSO (Recommended)**](#4-folder-operations-using-fso-recommended)
  - [**4.1 Create a folder (automatic parent creation)**](#41-create-a-folder-automatic-parent-creation)
  - [**4.2 Delete folder (including contents)**](#42-delete-folder-including-contents)
  - [**4.3 Copy entire folder**](#43-copy-entire-folder)
  - [**4.4 Move folder**](#44-move-folder)
  - [**4.5 Enumerate files in a folder**](#45-enumerate-files-in-a-folder)
  - [**4.6 Enumerate subfolders**](#46-enumerate-subfolders)
- [5. **Drive Interaction**](#5-drive-interaction)
  - [**5.1 List available drives**](#51-list-available-drives)
  - [**5.2 Get drive details**](#52-get-drive-details)
- [6. **Environment Variables and System Paths**](#6-environment-variables-and-system-paths)
  - [**6.1 Access environment variables**](#61-access-environment-variables)
  - [**6.2 Build system‑independent paths**](#62-build-systemindependent-paths)
- [7. **Recursively Traversing Folder Structures**](#7-recursively-traversing-folder-structures)
  - [**7.1 Recursive traversal example**](#71-recursive-traversal-example)
- [8. **Monitoring Folder Changes (Polling Pattern)**](#8-monitoring-folder-changes-polling-pattern)
  - [**8.1 Detect new files**](#81-detect-new-files)
- [9. **Robust Error Handling for Filesystem Work**](#9-robust-error-handling-for-filesystem-work)
  - [**Example Template**](#example-template)
- [10. **Advanced Examples**](#10-advanced-examples)
  - [**10.1 Search for files by extension across all subfolders**](#101-search-for-files-by-extension-across-all-subfolders)
  - [**10.2 Clean temporary files older than 7 days**](#102-clean-temporary-files-older-than-7-days)
- [**Advanced Folder Recursion Examples in VBA**](#advanced-folder-recursion-examples-in-vba)
- [1. **Recursive Listing of All Files (FSO Version)**](#1-recursive-listing-of-all-files-fso-version)
- [2. **Recursively Search for Files by Extension**](#2-recursively-search-for-files-by-extension)
- [3. **Recursive Folder Size Calculation**](#3-recursive-folder-size-calculation)
- [4. **Recursive Deletion of Old Files (e.g., older than N days)**](#4-recursive-deletion-of-old-files-eg-older-than-n-days)
- [5. **Generate a Full Directory Tree Structure (Indented)**](#5-generate-a-full-directory-tree-structure-indented)
- [6. **Native VBA DIR-Based Recursion (High Performance)**](#6-native-vba-dir-based-recursion-high-performance)
- [7. **Recursive Copy of an Entire Folder Tree**](#7-recursive-copy-of-an-entire-folder-tree)
- [8. **Build a Collection of Matching Files (Return Results)**](#8-build-a-collection-of-matching-files-return-results)

***

Below is a **complete, comprehensive, expert‑level guide** to **filesystem interaction in Excel VBA**, analogous in style and depth to the file‑handling guide you received earlier.  
This topic is stable and not time‑sensitive, so a web search is not required.

*(Directories, paths, folder structure, enumeration, environment variables, drives, permissions, copying/moving/creating/deleting, FSO and native VBA)*

This guide includes:

1.  **Fundamentals of filesystem interaction**
2.  **Path handling (native VBA)**
3.  **Folder operations (native VBA + FSO)**
4.  **Drive information**
5.  **Environment variables & system paths**
6.  **Advanced folder traversal (recursion)**
7.  **Folder monitoring patterns**
8.  **Robust error handling**

***

# 1. **Fundamentals of Filesystem Interaction**

VBA interacts with the filesystem using two main mechanisms:

### **A. Native VBA built‑ins**

`Dir`, `MkDir`, `RmDir`, `ChDir`, `CurDir`, `GetAttr`, etc.

### **B. FileSystemObject (FSO)**

Provided by `Scripting.FileSystemObject`, offering a modern, object‑based API.

You will often combine both, depending on performance and readability requirements.

***

# 2. **Path Handling in VBA**

## **2.1 Get current working directory**

```vba
Debug.Print CurDir$
```

## **2.2 Change working directory**

```vba
ChDir "C:\Data"
```

## **2.3 Extract components from a path**

### **Get file name from full path**

```vba
Dim p As String: p = "C:\Data\Reports\file.txt"
Debug.Print Dir(p)
```

### **Get folder portion of path**

```vba
Dim folder As String
folder = Left(p, InStrRev(p, "\") - 1)
```

### **Combine folder + filename safely**

```vba
Function CombinePath(folder As String, file As String) As String
    If Right(folder, 1) <> "\" Then folder = folder & "\"
    CombinePath = folder & file
End Function
```

***

# 3. **Folder Operations (Native VBA)**

## **3.1 Create a folder**

```vba
MkDir "C:\Data\NewFolder"
```

## **3.2 Remove a folder (only if empty)**

```vba
RmDir "C:\Data\NewFolder"
```

## **3.3 Check if a folder exists**

Native VBA cannot check folder existence directly; use `Dir`:

```vba
Function FolderExists(path As String) As Boolean
    FolderExists = (Dir(path, vbDirectory) <> "")
End Function
```

***

# 4. **Folder Operations Using FSO (Recommended)**

Initialize FSO:

```vba
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
```

***

## **4.1 Create a folder (automatic parent creation)**

```vba
fso.CreateFolder "C:\Data\NewProject"
```

***

## **4.2 Delete folder (including contents)**

```vba
fso.DeleteFolder "C:\Data\TempFolder", True
```

**Warning:** `True` means *force*, deleting read‑only contents.

***

## **4.3 Copy entire folder**

```vba
fso.CopyFolder "C:\Source", "C:\Backup\SourceCopy"
```

***

## **4.4 Move folder**

```vba
fso.MoveFolder "C:\Data\Old", "C:\Data\New"
```

***

## **4.5 Enumerate files in a folder**

```vba
Dim folder As Object, file As Object

Set folder = fso.GetFolder("C:\Data")

For Each file In folder.Files
    Debug.Print file.Name, file.DateLastModified, file.Size
Next file
```

***

## **4.6 Enumerate subfolders**

```vba
Dim subf As Object
For Each subf In folder.SubFolders
    Debug.Print subf.Path
Next subf
```

***

# 5. **Drive Interaction**

## **5.1 List available drives**

```vba
Dim drv As Object

For Each drv In fso.Drives
    Debug.Print drv.Path, drv.DriveType, drv.IsReady
Next drv
```

Drive types (numeric):

| Value | Meaning   |
| ----- | --------- |
| 0     | Unknown   |
| 1     | Removable |
| 2     | Fixed     |
| 3     | Network   |
| 4     | CD‑ROM    |
| 5     | RAM disk  |

***

## **5.2 Get drive details**

```vba
Dim d As Object
Set d = fso.GetDrive("C:")

Debug.Print d.TotalSize
Debug.Print d.FreeSpace
Debug.Print d.VolumeName
```

***

# 6. **Environment Variables and System Paths**

## **6.1 Access environment variables**

```vba
Debug.Print Environ("TEMP")
Debug.Print Environ("USERNAME")
Debug.Print Environ("USERPROFILE")
```

Common variables:

| Variable         | Meaning                 |
| ---------------- | ----------------------- |
| `%TEMP%`         | Temporary directory     |
| `%APPDATA%`      | Application data folder |
| `%PROGRAMFILES%` | Program Files           |
| `%HOMEPATH%`     | User home               |
| `%COMPUTERNAME%` | PC name                 |

***

## **6.2 Build system‑independent paths**

Example: Documents folder

```vba
Dim docs As String
docs = Environ("USERPROFILE") & "\Documents"
```

***

# 7. **Recursively Traversing Folder Structures**

This is essential for:

*   Searching for files
*   Batch processing
*   Indexing directory trees

***

## **7.1 Recursive traversal example**

```vba
Sub Traverse(path As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folder As Object, subf As Object, file As Object

    Set folder = fso.GetFolder(path)

    ' Files
    For Each file In folder.Files
        Debug.Print file.Path
    Next file

    ' Recurse
    For Each subf In folder.SubFolders
        Traverse subf.Path
    Next subf
End Sub
```

Usage:

```vba
Call Traverse("C:\Data")
```

***

# 8. **Monitoring Folder Changes (Polling Pattern)**

*(VBA cannot subscribe to OS events; use timed polling)*

## **8.1 Detect new files**

```vba
Static lastCount As Long

Sub CheckFolder()
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folder As Object

    Set folder = fso.GetFolder("C:\Watch")

    If folder.Files.Count <> lastCount Then
        Debug.Print "Change detected at " & Now
        lastCount = folder.Files.Count
    End If
End Sub
```

Run via `Application.OnTime`.

***

# 9. **Robust Error Handling for Filesystem Work**

## **Example Template**

```vba
Sub SafeFSO()
    On Error GoTo ErrHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    fso.CreateFolder "C:\Secure\Demo"

    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub
```

***

# 10. **Advanced Examples**

***

## **10.1 Search for files by extension across all subfolders**

```vba
Sub FindExcelFiles(path As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folder As Object, file As Object, subf As Object

    Set folder = fso.GetFolder(path)

    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "xlsx" Then
            Debug.Print file.Path
        End If
    Next file

    For Each subf In folder.SubFolders
        FindExcelFiles subf.Path
    Next subf
End Sub
```

***

## **10.2 Clean temporary files older than 7 days**

```vba
Sub CleanTemp()
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folder As Object, file As Object

    Set folder = fso.GetFolder(Environ("TEMP"))

    For Each file In folder.Files
        If Now - file.DateLastModified > 7 Then
            On Error Resume Next 'skip locked files
            fso.DeleteFile file.Path, True
        End If
    Next file
End Sub
```

***

Below is an **expanded, expert‑level set of folder‑recursion examples** for VBA.  
No external searches are needed — directory recursion in VBA is a stable, well‑documented topic.

The examples are designed to be **practical, progressive, and reusable**.

***

# **Advanced Folder Recursion Examples in VBA**

Recursion is essential when:

*   Processing deeply nested directories
*   Searching for files
*   Performing selective cleanup
*   Generating inventory reports
*   Applying actions to complex folder trees

We will use both **native VBA (`Dir`)** and **FileSystemObject (FSO)** versions because each has strengths.

***

# 1. **Recursive Listing of All Files (FSO Version)**

*(Most readable and recommended)*

```vba
Sub ListAllFiles(path As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folder As Object: Set folder = fso.GetFolder(path)
    Dim subf As Object, file As Object

    ' Files in current folder
    For Each file In folder.Files
        Debug.Print file.Path
    Next file

    ' Recurse into each subfolder
    For Each subf In folder.SubFolders
        ListAllFiles subf.Path
    Next subf
End Sub
```

Usage:

```vba
Call ListAllFiles("C:\Data")
```

***

# 2. **Recursively Search for Files by Extension**

```vba
Sub FindFilesByExtension(path As String, ext As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folder As Object: Set folder = fso.GetFolder(path)
    Dim subf As Object, file As Object

    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = LCase(ext) Then
            Debug.Print file.Path
        End If
    Next file

    For Each subf In folder.SubFolders
        FindFilesByExtension subf.Path, ext
    Next subf
End Sub
```

Usage:

```vba
Call FindFilesByExtension("C:\Projects", "xlsm")
```

***

# 3. **Recursive Folder Size Calculation**

```vba
Function FolderSize(path As String) As Double
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folder As Object: Set folder = fso.GetFolder(path)
    Dim file As Object, subf As Object
    Dim size As Double

    ' Add the size of each file
    For Each file In folder.Files
        size = size + file.Size
    Next file

    ' Add the size of each subfolder
    For Each subf In folder.SubFolders
        size = size + FolderSize(subf.Path)
    Next subf

    FolderSize = size
End Function
```

Usage:

```vba
Debug.Print "Folder size:", FolderSize("C:\Data") / 1024 / 1024 & " MB"
```

***

# 4. **Recursive Deletion of Old Files (e.g., older than N days)**

```vba
Sub DeleteOldFiles(path As String, daysOld As Double)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folder As Object: Set folder = fso.GetFolder(path)
    Dim file As Object, subf As Object

    For Each file In folder.Files
        If Now - file.DateLastModified > daysOld Then
            On Error Resume Next ' skip locked files
            fso.DeleteFile file.Path, True
        End If
    Next file

    For Each subf In folder.SubFolders
        DeleteOldFiles subf.Path, daysOld
    Next subf
End Sub
```

Usage:

```vba
' Delete all files older than 30 days
Call DeleteOldFiles("C:\Logs", 30)
```

***

# 5. **Generate a Full Directory Tree Structure (Indented)**

*(Useful for reports or export)*

```vba
Sub PrintTree(path As String, Optional level As Long = 0)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folder As Object: Set folder = fso.GetFolder(path)
    Dim subf As Object, file As Object
    Dim indent As String

    indent = String(level * 4, " ")

    Debug.Print indent & "[" & folder.Name & "]"

    ' Files
    For Each file In folder.Files
        Debug.Print indent & "    " & file.Name
    Next file

    ' Subfolders
    For Each subf In folder.SubFolders
        PrintTree subf.Path, level + 1
    Next subf
End Sub
```

Usage:

```vba
Call PrintTree("C:\Projects")
```

Output example:

    [Projects]
        file1.txt
        file2.log
        [Archive]
            old1.txt
            old2.txt
        [Scripts]
            run.vbs
            utils.vba

***

# 6. **Native VBA DIR-Based Recursion (High Performance)**

DIR recursion is faster than FSO but less readable.

```vba
Sub RecDir(path As String)
    Dim subpath As String
    Dim file As String

    ' Ensure trailing slash
    If Right(path, 1) <> "\" Then path = path & "\"

    ' Files
    file = Dir(path & "*.*", vbNormal Or vbHidden Or vbReadOnly)
    Do While file <> ""
        Debug.Print path & file
        file = Dir()
    Loop

    ' Subfolders
    subpath = Dir(path & "*", vbDirectory Or vbHidden Or vbReadOnly)
    Do While subpath <> ""
        If subpath <> "." And subpath <> ".." Then
            If (GetAttr(path & subpath) And vbDirectory) = vbDirectory Then
                RecDir path & subpath
            End If
        End If
        subpath = Dir()
    Loop
End Sub
```

Usage:

```vba
Call RecDir("C:\Data")
```

***

# 7. **Recursive Copy of an Entire Folder Tree**

*(FSO‑based clone operation)*

```vba
Sub RecursiveCopy(src As String, dst As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folder As Object: Set folder = fso.GetFolder(src)
    Dim subf As Object, file As Object

    If Not fso.FolderExists(dst) Then
        fso.CreateFolder dst
    End If

    ' Copy files
    For Each file In folder.Files
        fso.CopyFile file.Path, dst & "\" & file.Name
    Next file

    ' Recurse into subfolders
    For Each subf In folder.SubFolders
        RecursiveCopy subf.Path, dst & "\" & subf.Name
    Next subf
End Sub
```

Usage:

```vba
RecursiveCopy "C:\Source", "C:\Backup\Source"
```

***

# 8. **Build a Collection of Matching Files (Return Results)**

```vba
Function FindFiles(path As String, pattern As String) As Collection
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim col As New Collection
    Call SearchRec(path, pattern, col)
    Set FindFiles = col
End Function

Private Sub SearchRec(path As String, pattern As String, col As Collection)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folder As Object: Set folder = fso.GetFolder(path)
    Dim file As Object, subf As Object

    For Each file In folder.Files
        If LCase(file.Name) Like LCase(pattern) Then
            col.Add file.Path
        End If
    Next file

    For Each subf In folder.SubFolders
        SearchRec subf.Path, pattern, col
    Next subf
End Sub
```

Usage:

```vba
Dim results As Collection
Set results = FindFiles("C:\Data", "*.log")

Dim item As Variant
For Each item In results
    Debug.Print item
Next item
```

---

[DOC MOC](./doc-00_MOC.md)