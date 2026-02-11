# Debug



## Immediate Window

Select **View -> Immediate Window** to activate. Any valid expression can be evaluated in the Immediate Window.

In the window it is possible to use the _Print method_ as follows:

> print [items] [;]

otherwise, the _print_ command can be replaced by a question mark (_?_) as follows:

> ? [item]

## MessageBox

VBA _MsgBox_ is a function that generates a dialog window.

The syntax is the following:

> MsgBox(prompt[, buttons] [, title] [, helpfile, context])

where:

- _prompt_ is the message to be visualized in the box.

- _buttons_ are the **constants** of the argument that allow to choose the type of message box:

  | Constant | Value | Description |
  | -------- | :---: | ----------- |
  |vbOKOnly|0|_ok_ button|
  |VbOKCancel|1|_ok_ and _cancel_ buttons|
  |VbAbortRetryIgnore|2|_cancel_, _retry_ and _ignore_ buttons|
  |VbYesNoCancel|3|_yes_, _no_ and _cancel_ buttons|
  |VbYesNo|4| _yes_ and _no_ buttons |
  |VbRetryCancel|5| _retry_ and _cancel_ buttons |
  |VbCritical|16| critical message |
  |VbQuestion|32| reboot request |
  |VbExclamation|48| advising message |
  |VbInformation|64| information message |

  When a **constant** is used, the instruction can be assigned to a variable (that must be declared as String or  Variant (Dim) depending on the output of the function). The instruction will be

  > Dim variable
  >
  > variable = MsgBox([message], [constant])
  
- The length of a message is 1024 characters. Special commands can be used

  | Command | Alternative Command | What It Does                          |
  | ------- | ------------------- | ------------------------------------- |
  | vbCr    | Chr(13)             | go to next line                       |
  | vbLf    | Chr(10)             | empty line                            |
  | vbCrLf  | Chr(13) + Chr(10)   | go to next line and add an empty line |

  Below an example on how to write a multi-line message:

  ```vb
  Sub Messaggio()
  	Dim Lem As String
  	Lem = Lem & "Questo è un esempio :" & vbLf & vbLf
  	Lem = Lem & "L'Autore è:" & Chr(13) & vbLf
  	Lem = Lem & "Pinco Pallino" & vbLf
  	Lem = Lem & "Questo Programma" & vbLf
  	Lem = Lem & "è tutelato dai" & vbLf & vbLf
  	Lem = Lem & "diritti d'Autore !!" & vbLf
  	Lem = Lem & "(non esiste, è falso)" & vbLf & vbLf
  	Lem = Lem & "del Codice VBa di questo messaggio" & vbLf
  	Lem = Lem & "si raccomanda di fare tutte" & vbLf
  	Lem = Lem & "le variazioni che volete." & vbLf
  	Lem = Lem & "l'Autore lo concede (?!?!?)" & vbLf
  	Lem = Lem & "Ha......Ha......Ha"
  	MsgBox Lem
  End Sub
  ```

Some examples of MsgBox messages:

```vb
MsgBox "You're in the wrong sheet!" & vbNewLine & "Please switch to the sheet: " & sheetName, vbCritical

MsgBox "OK!" & vbCrLf, vbInformation

```

Operations when closing a file (with _Cases_):

```vb
Sub Prova4()
	Dim iRisposta As Integer
	iRisposta = MsgBox("STAI PER USCIRE, VUOI SALVARE IL FILE ???", vbYesNoCancel)
	Select Case iRisposta 'impostiamo il Select Case con riferimento al messaggio restituito dalla variabile iRisposta
	Case vbYes 'se risponderemo "Si" :
		ThisWorkbook.Save 'salveremo il file e
		Application.Quit 'chiuderemo cartella ed Excel
	Case vbCancel 'se sceglieremo "Annulla":
		Exit Sub 'usciremo dalla routine
	Case vbNo 'se sceglieremo "No":
		Application.Quit 'chiuderemo cartella ed Excel senza salvare
	Case Else
		End Select
End Sub
```

Set a predefined value (_vbNo_) to not accomplish an action:

> Cancel = (MsgBox("Sicuro di voler chiudere la finestra ?", vbYesNo) = vbNo)

## Print

The _Print method_ sends output to the immediate window whenever the _Debug object_ prefix is included:

> Debug.Print [items] [;]

This string can be written in any row of the code.
