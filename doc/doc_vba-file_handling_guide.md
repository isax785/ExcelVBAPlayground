# **Complete Guide to File Handling in Excel VBA**

- [**Complete Guide to File Handling in Excel VBA**](#complete-guide-to-file-handling-in-excel-vba)
- [1. **Basic File System Functions in VBA**](#1-basic-file-system-functions-in-vba)
  - [**1.1 DIR – Check if a file/folder exists**](#11-dir--check-if-a-filefolder-exists)
    - [**Check if a file exists**](#check-if-a-file-exists)
    - [**List all files in a folder**](#list-all-files-in-a-folder)
  - [**1.2 FileLen – Get file size**](#12-filelen--get-file-size)
  - [**1.3 GetAttr – Get file/folder attributes**](#13-getattr--get-filefolder-attributes)
- [2. **Reading and Writing TEXT Files**](#2-reading-and-writing-text-files)
  - [**2.1 Write Text (Output mode)**](#21-write-text-output-mode)
  - [**2.2 Append Text**](#22-append-text)
  - [**2.3 Read Text File Line‑by‑Line (Input mode)**](#23-read-text-file-linebyline-input-mode)
- [3. **Reading and Writing BINARY Files**](#3-reading-and-writing-binary-files)
  - [**3.1 Read Entire Binary File to Byte Array**](#31-read-entire-binary-file-to-byte-array)
  - [**3.2 Write Binary File from Byte Array**](#32-write-binary-file-from-byte-array)
- [4. **FileSystemObject (FSO)**](#4-filesystemobject-fso)
- [5. **File Operations Using FSO**](#5-file-operations-using-fso)
  - [**5.1 Check if file exists**](#51-check-if-file-exists)
  - [**5.2 Create a new file**](#52-create-a-new-file)
  - [**5.3 Read a text file with FSO**](#53-read-a-text-file-with-fso)
  - [**5.4 Copy, delete, rename**](#54-copy-delete-rename)
- [6. **Working with Folders**](#6-working-with-folders)
  - [**6.1 Check if folder exists**](#61-check-if-folder-exists)
  - [**6.2 Create folder**](#62-create-folder)
  - [**6.3 Loop through files**](#63-loop-through-files)
- [7. **Error Handling Best Practices**](#7-error-handling-best-practices)
  - [**7.1 Safe‑open with error capture**](#71-safeopen-with-error-capture)
- [8. **Advanced Examples**](#8-advanced-examples)
  - [**8.1 Write an array to a CSV file**](#81-write-an-array-to-a-csv-file)
  - [**8.2 Read a CSV into a dynamic array**](#82-read-a-csv-into-a-dynamic-array)
  - [**8.3 Backup all Excel files in a folder**](#83-backup-all-excel-files-in-a-folder)

***

Below is a **full, comprehensive, expert‑level guide** to **file handling in Excel VBA**, complete with best‑practice explanations and meaningful real‑world examples.  
No search is required because the topic is stable, well‑defined, and not time‑sensitive.


This guide covers:

1.  **File System Basics (Dir, FileLen, GetAttr)**
2.  **Reading/Writing Text Files**
3.  **Working with Binary Files**
4.  **Using the FileSystemObject (FSO)**
5.  **Creating, copying, deleting, and renaming files**
6.  **Working with folders**
7.  **Handling errors & safe‑coding practices**
8.  **Advanced examples**

***

# 1. **Basic File System Functions in VBA**

## **1.1 DIR – Check if a file/folder exists**

### **Check if a file exists**

```vba
Dim path As String
path = "C:\Data\report.txt"

If Dir(path) <> "" Then
    MsgBox "File exists"
Else
    MsgBox "File does not exist"
End If
```

### **List all files in a folder**

```vba
Dim file As String
file = Dir("C:\Data\*.txt")

Do While file <> ""
    Debug.Print file
    file = Dir()
Loop
```

***

## **1.2 FileLen – Get file size**

```vba
Dim size As Long
size = FileLen("C:\Data\report.txt")
Debug.Print "Size = " & size & " bytes"
```

***

## **1.3 GetAttr – Get file/folder attributes**

```vba
Dim att As Long
att = GetAttr("C:\Data\report.txt")

If att And vbReadOnly Then Debug.Print "Read-only"
If att And vbHidden Then Debug.Print "Hidden"
If att And vbDirectory Then Debug.Print "Directory"
```

***

# 2. **Reading and Writing TEXT Files**

Excel VBA supports **three modes**:

| Mode       | Purpose                   |
| ---------- | ------------------------- |
| **Input**  | Read text                 |
| **Output** | Overwrite + write text    |
| **Append** | Add text to existing file |

***

## **2.1 Write Text (Output mode)**

```vba
Dim f As Integer
f = FreeFile

Open "C:\Data\log.txt" For Output As #f
Print #f, "This is line 1"
Print #f, "This is line 2"
Close #f
```

***

## **2.2 Append Text**

```vba
Dim f As Integer
f = FreeFile

Open "C:\Data\log.txt" For Append As #f
Print #f, "New entry added at " & Now
Close #f
```

***

## **2.3 Read Text File Line‑by‑Line (Input mode)**

```vba
Dim f As Integer, line As String
f = FreeFile

Open "C:\Data\log.txt" For Input As #f

Do Until EOF(f)
    Line Input #f, line
    Debug.Print line
Loop

Close #f
```

***

# 3. **Reading and Writing BINARY Files**

Use for images, PDFs, or any non-text content.

***

## **3.1 Read Entire Binary File to Byte Array**

```vba
Dim f As Integer
Dim bytes() As Byte

f = FreeFile

Open "C:\Data\image.jpg" For Binary As #f
ReDim bytes(LOF(f) - 1)
Get #f, , bytes
Close #f
```

***

## **3.2 Write Binary File from Byte Array**

```vba
Dim f As Integer
Dim bytes() As Byte
' (bytes array already populated)

f = FreeFile
Open "C:\Data\copy.jpg" For Binary As #f
Put #f, , bytes
Close #f
```

***

# 4. **FileSystemObject (FSO)**

A more modern and powerful abstraction.

```vba
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
```

***

# 5. **File Operations Using FSO**

## **5.1 Check if file exists**

```vba
If fso.FileExists("C:\Data\report.txt") Then
    MsgBox "Exists"
End If
```

***

## **5.2 Create a new file**

```vba
Dim file As Object
Set file = fso.CreateTextFile("C:\Data\newfile.txt", True)
file.WriteLine "Hello world"
file.Close
```

***

## **5.3 Read a text file with FSO**

```vba
Dim ts As Object
Set ts = fso.OpenTextFile("C:\Data\report.txt", 1)

Do Until ts.AtEndOfStream
    Debug.Print ts.ReadLine
Loop

ts.Close
```

***

## **5.4 Copy, delete, rename**

```vba
fso.CopyFile "C:\Data\a.txt", "C:\Backup\a.txt"
fso.DeleteFile "C:\Data\b.txt"
fso.MoveFile "C:\Data\oldname.txt", "C:\Data\newname.txt"
```

***

# 6. **Working with Folders**

## **6.1 Check if folder exists**

```vba
If fso.FolderExists("C:\Data") Then
    MsgBox "Folder found"
End If
```

***

## **6.2 Create folder**

```vba
fso.CreateFolder "C:\NewFolder"
```

***

## **6.3 Loop through files**

```vba
Dim folder As Object, file As Object
Set folder = fso.GetFolder("C:\Data")

For Each file In folder.Files
    Debug.Print file.Name, file.Size
Next file
```

***

# 7. **Error Handling Best Practices**

## **7.1 Safe‑open with error capture**

```vba
On Error GoTo ErrHandler

Dim f As Integer
f = FreeFile
Open "C:\Data\data.txt" For Input As #f

' Read...

Close #f
Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description
    If f > 0 Then Close #f
```

***

# 8. **Advanced Examples**

***

## **8.1 Write an array to a CSV file**

```vba
Sub WriteCsv()
    Dim f As Integer, r As Long, c As Long
    Dim path As String: path = "C:\Data\export.csv"

    f = FreeFile
    Open path For Output As #f

    For r = 1 To 10
        Dim line As String: line = ""
        For c = 1 To 5
            line = line & Cells(r, c).Value & IIf(c < 5, ",", "")
        Next c
        Print #f, line
    Next r

    Close #f
End Sub
```

***

## **8.2 Read a CSV into a dynamic array**

```vba
Function LoadCsv(path As String) As Variant
    Dim f As Integer, txt As String
    f = FreeFile

    Open path For Input As #f
    txt = Input(LOF(f), #f)
    Close #f

    Dim rows() As String, lines() As String
    lines = Split(txt, vbCrLf)

    Dim arr() As Variant
    ReDim arr(0 To UBound(lines), 0 To 10)

    Dim r As Long
    For r = 0 To UBound(lines)
        rows = Split(lines(r), ",")
        Dim c As Long
        For c = 0 To UBound(rows)
            arr(r, c) = rows(c)
        Next c
    Next r

    LoadCsv = arr
End Function
```

***

## **8.3 Backup all Excel files in a folder**

```vba
Sub BackupExcelFiles()
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim src As Object: Set src = fso.GetFolder("C:\Data")
    
    Dim file As Object, backupFolder As String
    backupFolder = "C:\Backup\"

    If Not fso.FolderExists(backupFolder) Then
        fso.CreateFolder backupFolder
    End If

    For Each file In src.Files
        If LCase(fso.GetExtensionName(file.Name)) = "xlsx" Then
            fso.CopyFile file.Path, backupFolder & file.Name
        End If
    Next file
End Sub
```

---

[DOC MOC](./doc-00_MOC.md)