# Sub and Function

Details [here](../doc/doc-vba%20-%20sub%20and%20functions.md).

> Defaults (for both `Sub` and `Function`):  
> - `Public`
> - `ByRef`

- [ ] call a sub
- [ ] calla function

## Function

> Syntax:  
> *`[Public | Private | Friend] [ Static ] Function name [ ( arglist ) ] [ As type ] [ statements ] [ name = expression ] [ Exit Function ] [ statements ] [name = expression ] End Function`*
> `arglist` syntax:  
> *`[ Optional ] [ ByVal | ByRef ] [ ParamArray ] varname [ ( ) ] [ As type ] [ = defaultvalue ]`*

The name of the function corresponds to the function output and must be declared in it's first statement.

```vb
Function calc_sum(a as Integer, b as Integer) as Integer
    calc_sum = a + b
End Function
```

## Sub

> Syntax:  
> *`[ Private | Public | Friend ] [ Static ] Sub name [ ( arglist ) ] [ statements ] [ Exit Sub ] [ statements ] End Sub`*
> `arglist` syntax:  
> *`[ Optional ] [ ByVal | ByRef ] [ ParamArray ] varname [ ( ) ] [ As type ] [ = defaultvalue ]`*


**Event Handlers** are not stored in your typical module location. They are actually stored inside either your Workbook or Worksheet object. To get to the "coding area" of either your workbook or worksheet, you simply double-click **ThisWorkbook** or the sheet name (respectively) within your desired VBA Project hierarchy tree (within the **Project Window** of your Visual Basic Editor).


---

[MOC](./tbx%20-%2000%20MOC.md)