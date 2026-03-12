Tags: #guidelines #tools #toolbox

---

# Toolbox Guidelines

- [[#Table Example]]

---

Guidelines for the development of toolboxes, i.e. ready to use **reference files** with the following main purpose: **speed-up, ease, and reduce energy consumption of developments**.

- use *markdown tables* as command reference;
- split tables with the proper granularity:
	- global split: divide topics into separated tables;
	- internal split: organize each topic to speed-up finding the needed reference.
- use *snippets*: 
	- only whenever there i no possibility to store the code into the markdown table;
	- must be shorter as possibile, very focused on the single operation/action. 
	- add a few words for explanation  only if strictly needed, better if in the form of code comments.
- use titles to be used in a table of contents reference;
- if the toolbox is too long, split it into multiple files;
- filename **prefix**: `tbx-[area_ref] - [topic]`. Examples: 
	- `tbx-py - core syntax`;
	- sub-area: `tbx-py-re - pattern formats`.
- tables:
    - to ease the comprehension of the action, report also the returned object, e.g. parse string (->`list[str]`);
    - when providing the **syntax** for the action (i.e. general statement instead of specific code):
    	- write code in *italic*, e.g. *`def [func_name]`*;
     	- parameters are to be enclosed between symbols depending on the language to avoid superposition with the language syntax, e.g. Python -> `<...>`, C++ -> `[...]`, VB -> `[...]`.

## Table Example

**Some Topic**

| Action                  | Code        |
| ----------------------- | ----------- |
| whatever you want to do | `a = b + c` |
| **Subtopic**            |             |
| another command         | `d = f(e)`  |
| *Sub-subtopic*          |             |
| more and more           | `f(a + d)`  |

*Subtopic*

> a very compact snippet.

```python
def f(a:int):
	...
	a *= 10
	...
	return a
```
