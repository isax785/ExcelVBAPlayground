# 💻 **Shell Prompting with Excel VBA — Complete Expert Guide**

- [💻 **Shell Prompting with Excel VBA — Complete Expert Guide**](#-shell-prompting-with-excel-vba--complete-expert-guide)
- [1️⃣ What Is Shell Prompting in VBA?](#1️⃣-what-is-shell-prompting-in-vba)
- [2️⃣ Core Methods for Shell Execution in VBA](#2️⃣-core-methods-for-shell-execution-in-vba)
  - [**A. The `Shell` Function (basic method)**](#a-the-shell-function-basic-method)
  - [**B. `WScript.Shell.Run` (better control)**](#b-wscriptshellrun-better-control)
  - [**C. `WScript.Shell.Exec` (best method)**](#c-wscriptshellexec-best-method)
  - [**D. `ShellExecute` API (advanced use)**](#d-shellexecute-api-advanced-use)
- [3️⃣ Asynchronous vs Synchronous Execution](#3️⃣-asynchronous-vs-synchronous-execution)
    - [**Asynchronous**](#asynchronous)
    - [**Synchronous**](#synchronous)
- [4️⃣ Capturing Command Output (Console Text)](#4️⃣-capturing-command-output-console-text)
    - [Full example:](#full-example)
- [5️⃣ Passing Parameters \& Correct Quoting](#5️⃣-passing-parameters--correct-quoting)
    - [Example: path containing spaces](#example-path-containing-spaces)
    - [Example: Run PowerShell with parameters](#example-run-powershell-with-parameters)
- [6️⃣ Waiting For External Programs to Finish](#6️⃣-waiting-for-external-programs-to-finish)
- [7️⃣ Running PowerShell from Excel VBA](#7️⃣-running-powershell-from-excel-vba)
    - [Example: Running a PowerShell script file](#example-running-a-powershell-script-file)
    - [Example: Calling REST API via PowerShell from Excel](#example-calling-rest-api-via-powershell-from-excel)
- [8️⃣ Most Common Real‑World Applications](#8️⃣-most-common-realworld-applications)
  - [✔ 1. Running Batch Files from Excel](#-1-running-batch-files-from-excel)
  - [✔ 2. Automating System Commands](#-2-automating-system-commands)
  - [✔ 3. Calling External Programs](#-3-calling-external-programs)
  - [✔ 4. Running PowerShell for advanced automation](#-4-running-powershell-for-advanced-automation)
  - [✔ 5. Ping, network checks, diagnostics](#-5-ping-network-checks-diagnostics)
  - [✔ 6. Git, Python, Node, Java automation](#-6-git-python-node-java-automation)
  - [✔ 7. Reading system information](#-7-reading-system-information)
  - [✔ 8. Automating software builds](#-8-automating-software-builds)
- [9️⃣ Full Reusable VBA Shell Toolkit](#9️⃣-full-reusable-vba-shell-toolkit)
- [🔟 Summary (What You Now Know)](#-summary-what-you-now-know)


***

Below is a **full, expert‑level, comprehensive overview of Shell Prompting with Excel VBA**, including **core concepts**, **best practices**, **common applications**, and **fully working code examples**.

Running shell commands from Excel VBA allows you to automate Windows tasks, interact with external programs, execute batch/powershell scripts, and capture command output. This capability turns Excel into a powerful automation hub.

This guide covers:

1.  **What “Shell Prompting” Means in VBA**
2.  **Core Methods: `Shell`, `WScript.Shell`, `Exec`, `Run`**
3.  **Asynchronous vs Synchronous execution**
4.  **Capturing command output (console text)**
5.  **Passing parameters, quoting rules**
6.  **Waiting for a program to finish**
7.  **Launching PowerShell from VBA**
8.  **Most common real‑world applications**

***

# 1️⃣ What Is Shell Prompting in VBA?

“Shell prompting” refers to VBA code that:

*   Executes **Windows commands** (cmd.exe)
*   Runs **external programs or executables**
*   Launches **batch files (.bat)** or **PowerShell scripts (.ps1)**
*   Issues system operations like file copy, ping, curl, robocopy…

Excel uses **the Windows Shell** to trigger these operations.

***

# 2️⃣ Core Methods for Shell Execution in VBA

There are four key approaches.

***

## **A. The `Shell` Function (basic method)**

✔ Simple  
✔ Fast  
✖ No output capture  
✖ Asynchronous (doesn't wait)

```vba
Shell "cmd.exe /c dir C:\", vbNormalFocus
```

***

## **B. `WScript.Shell.Run` (better control)**

✔ Can run synchronously  
✔ Return exit code  
✖ Cannot capture command output

```vba
Dim sh As Object
Set sh = CreateObject("WScript.Shell")

Dim exitCode As Long
exitCode = sh.Run("cmd.exe /c echo Hello", 0, True)
```

`True` = VBA waits until the command finishes.

***

## **C. `WScript.Shell.Exec` (best method)**

✔ Capture console output  
✔ Read real‑time output  
✔ Check exit code

```vba
Dim sh As Object, exec As Object, output As String
Set sh = CreateObject("WScript.Shell")

Set exec = sh.Exec("cmd.exe /c ipconfig")

Do While Not exec.StdOut.AtEndOfStream
    output = output & exec.StdOut.ReadLine & vbCrLf
Loop

MsgBox output
```

This is the *professional* technique used to automate processes.

***

## **D. `ShellExecute` API (advanced use)**

✔ Best for opening files/programs  
✔ Supports verbs (“open”, “print”, “runas”)  
✖ Cannot capture output

Often used to launch programs with admin privileges.

***

# 3️⃣ Asynchronous vs Synchronous Execution

### **Asynchronous**

VBA continues running without waiting.

```vba
Shell "notepad.exe", vbNormalFocus
```

### **Synchronous**

Program must finish before VBA continues.

```vba
sh.Run "mybatch.bat", 0, True
```

***

# 4️⃣ Capturing Command Output (Console Text)

This is the **#1 most requested** shell functionality.

### Full example:

```vba
Function RunCMD(cmd As String) As String
    Dim sh As Object: Set sh = CreateObject("WScript.Shell")
    Dim exec As Object: Set exec = sh.Exec("cmd.exe /c " & cmd)

    Dim output As String
    Do Until exec.StdOut.AtEndOfStream
        output = output & exec.StdOut.ReadLine & vbCrLf
    Loop

    RunCMD = output
End Function
```

Usage:

```vba
Sub Test()
    Debug.Print RunCMD("ping 8.8.8.8")
End Sub
```

***

# 5️⃣ Passing Parameters & Correct Quoting

Windows shell is very sensitive to quotes.

### Example: path containing spaces

```vba
Shell "cmd.exe /c ""C:\Program Files\MyApp\run.exe"" -arg1 -arg2"
```

### Example: Run PowerShell with parameters

```vba
sh.Run "powershell -NoProfile -Command ""Get-ChildItem 'C:\Temp'""", 0, True
```

***

# 6️⃣ Waiting For External Programs to Finish

When you need Excel to pause:

```vba
Dim sh As Object: Set sh = CreateObject("WScript.Shell")
sh.Run "myScript.bat", 0, True  'True = Wait
```

Or using a busy-loop:

```vba
Dim PID As Long
PID = Shell("notepad.exe")
Do While ProcessExists(PID)
    DoEvents
Loop
```

***

# 7️⃣ Running PowerShell from Excel VBA

Powerful for:  
✔ API calls  
✔ File system automation  
✔ JSON parsing  
✔ Git operations  
✔ Docker, WSL, Azure CLI, etc.

### Example: Running a PowerShell script file

```vba
sh.Run "powershell -ExecutionPolicy Bypass -File ""C:\Scripts\myscript.ps1""", 1, True
```

### Example: Calling REST API via PowerShell from Excel

```vba
Dim cmd As String
cmd = "powershell -command ""Invoke-WebRequest 'https://api.ipify.org' -UseBasicParsing"""
Debug.Print RunCMD(cmd)
```

***

# 8️⃣ Most Common Real‑World Applications

## ✔ 1. Running Batch Files from Excel

Most common in enterprise automation.

```vba
Shell "cmd.exe /c C:\Scripts\deploy.bat"
```

***

## ✔ 2. Automating System Commands

*   `copy`
*   `move`
*   `robocopy`
*   `taskkill`
*   `shutdown`

```vba
Call RunCMD("robocopy C:\Source C:\Dest *.txt")
```

***

## ✔ 3. Calling External Programs

Used for CAD tools, simulations, compilers, utilities.

```vba
Shell """C:\Program Files\7-Zip\7z.exe"" x file.zip", vbHide
```

***

## ✔ 4. Running PowerShell for advanced automation

Example: create directory structure:

```vba
sh.Run "powershell -command ""New-Item -ItemType Directory -Path 'C:\Temp\Test'""", 0, True
```

***

## ✔ 5. Ping, network checks, diagnostics

```vba
Debug.Print RunCMD("ping google.com")
```

***

## ✔ 6. Git, Python, Node, Java automation

Example: run Python script from Excel:

```vba
Shell "cmd.exe /c python C:\Scripts\process_data.py"
```

***

## ✔ 7. Reading system information

```vba
Debug.Print RunCMD("systeminfo")
```

***

## ✔ 8. Automating software builds

Excel orchestrates:

*   Git pull
*   Build tools
*   Packaging
*   Deployment scripts

***

# 9️⃣ Full Reusable VBA Shell Toolkit

Below is a full module you can copy into Excel.

```vba
'==========================================
' SHELL TOOLKIT FOR VBA
'==========================================

' Run command and WAIT
Function ShellWait(cmd As String) As Long
    Dim sh As Object: Set sh = CreateObject("WScript.Shell")
    ShellWait = sh.Run("cmd.exe /c " & cmd, 0, True)
End Function

' Run command and CAPTURE output
Function ShellCapture(cmd As String) As String
    Dim sh As Object: Set sh = CreateObject("WScript.Shell")
    Dim exec As Object: Set exec = sh.Exec("cmd.exe /c " & cmd)
    
    Dim s As String
    Do Until exec.StdOut.AtEndOfStream
        s = s & exec.StdOut.ReadLine & vbCrLf
    Loop
    
    ShellCapture = s
End Function

' Run PowerShell
Function PowerShellCapture(ps As String) As String
    PowerShellCapture = ShellCapture("powershell -NoProfile -Command """ & ps & """")
End Function
```

***

# 🔟 Summary (What You Now Know)

You now have a full professional toolkit for:

✔ Running shell commands  
✔ Running batch files  
✔ Running PowerShell commands  
✔ Waiting for external processes  
✔ Capturing command output  
✔ Automating system tasks  
✔ Integrating Excel with external programs

This is the **complete, expert-level overview** of shell prompting in Excel VBA.


---

[DOC MOC](./doc-00_MOC.md)