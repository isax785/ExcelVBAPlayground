# Excel Toolbox

## `VLOOKUP` 

```excel
=VLOOKUP(B93,$C$2:$C$254,1,FALSE)
```

## Split Strings

[Source](https://www.ablebits.com/office-addins-blog/split-text-string-excel/)

> If you want to find an actual question mark or asterisk, **type a tilde (~) before the character**.

![[Pasted image 20240307110318.png]]

To extract the **item name** (all characters before the 1st hyphen), insert the following formula in B2, and then copy it down the column: 

```
=LEFT(A2, SEARCH("-",A2,1)-1)
```   

To extract the **color** (all characters between the 1st and 2nd hyphens), enter the following formula in C2, and then copy it down to other cells: 

```
=MID(A2, SEARCH("-",A2) + 1, SEARCH("-",A2,SEARCH("-",A2)+1) - SEARCH("-",A2) - 1)
``` 

To extract the **size** (all characters after the 3rd hyphen), enter the following formula in D2: 

```
=RIGHT(A2,LEN(A2) - SEARCH("-", A2, SEARCH("-", A2) + 1))
```
 
In a similar fashion, you can split column by any other character. All you have to do is to replace "-" with the required delimiter, for example **space** (" "), **comma** (","), **slash** ("/"), **colon** (";"), **semicolon** (";"), and so on.

## Quadratic Regression

`=LINEST(AG8:AG16,AK8:AK16^{1,2},1,1)`

![[Pasted image 20240805100906.png]]

# Add Arrow Inside a Cell

![add arrow to cell in excel 2013](https://cdn4syt-solveyourtech.netdna-ssl.com/wp-content/uploads/2018/09/how-insert-arrow-excel-2013-4.jpg)

