# Dynamic Arrays Filter and Lambda

- [Dynamic Arrays Filter and Lambda](#dynamic-arrays-filter-and-lambda)
- [1. Using `FILTER` with Dynamic Arrays](#1-using-filter-with-dynamic-arrays)
  - [1.1 Basic Syntax](#11-basic-syntax)
  - [1.2 Single Condition Example](#12-single-condition-example)
  - [1.3 Multiple Conditions (AND logic)](#13-multiple-conditions-and-logic)
  - [1.4 Multiple Conditions (OR logic)](#14-multiple-conditions-or-logic)
  - [1.5 Filtering by a List (Dynamic OR)](#15-filtering-by-a-list-dynamic-or)
  - [1.6 Dynamic Array Behavior](#16-dynamic-array-behavior)
- [2. Using `LAMBDA` Inside Array Formulas](#2-using-lambda-inside-array-formulas)
  - [2.1 Basic Inline LAMBDA Example](#21-basic-inline-lambda-example)
  - [2.2 Using LET + LAMBDA in Arrays](#22-using-let--lambda-in-arrays)
  - [2.3 Name Manager LAMBDA: reusable custom function](#23-name-manager-lambda-reusable-custom-function)
    - [Step 1 — Create a reusable function](#step-1--create-a-reusable-function)
    - [Step 2 — Use it](#step-2--use-it)
- [3. Array-Oriented LAMBDA Examples](#3-array-oriented-lambda-examples)
  - [3.1 Dot Product Function](#31-dot-product-function)
  - [3.2 Apply a Function Row‑wise or Column‑wise (MAP / BYROW / BYCOL)](#32-apply-a-function-rowwise-or-columnwise-map--byrow--bycol)
    - [Example: Row sum using BYROW](#example-row-sum-using-byrow)
    - [Example: Apply a transformation to two arrays using MAP](#example-apply-a-transformation-to-two-arrays-using-map)
  - [3.3 Build a Flexible Filter Using LAMBDA](#33-build-a-flexible-filter-using-lambda)
- [4. Example: Multi-condition Dynamic Filter with LAMBDA and MAP](#4-example-multi-condition-dynamic-filter-with-lambda-and-map)
- [If you want, I can also provide:](#if-you-want-i-can-also-provide)

---

Below is a structured, Excel-expert–level explanation of:

1.  **How to use `FILTER` with dynamic arrays**
2.  **How to use `LAMBDA` functions inside array formulas**

No external search is required because this is technical documentation.

***

# 1. Using `FILTER` with Dynamic Arrays

`FILTER` is one of the core dynamic array functions in modern Excel (Microsoft 365 / Excel 2021+).  
It extracts rows or columns based on logical conditions. The result **spills** automatically into adjacent cells.

***

## 1.1 Basic Syntax

```excel
FILTER(array, include, [if_empty])
```

*   **array** → the range you want to return filtered rows from
*   **include** → a Boolean array (TRUE/FALSE) of the **same height or width** as `array`
*   **if\_empty** → optional text or value returned if no rows match

***

## 1.2 Single Condition Example

**Data**

*   Table `A2:C10` with columns: `Category`, `Month`, `Revenue`
*   You want all rows where Category = `"HVAC"`

**Formula**

```excel
=FILTER(A2:C10, A2:A10="HVAC", "No matching category")
```

The returned result spills down automatically.

***

## 1.3 Multiple Conditions (AND logic)

To filter by Category *and* Month:

```excel
=FILTER(A2:C10, (A2:A10="HVAC") * (B2:B10="Jan"), "No results")
```

*   Multiplication `*` performs element-wise AND
*   TRUE evaluates to 1, FALSE to 0

***

## 1.4 Multiple Conditions (OR logic)

Return rows where Month is `"Jan"` **OR** `"Feb"`:

```excel
=FILTER(A2:C10, (B2:B10="Jan") + (B2:B10="Feb"), "No results")
```

*   Addition `+` performs element-wise OR
*   TRUE+TRUE = 2, TRUE+FALSE = 1, both treated as TRUE

***

## 1.5 Filtering by a List (Dynamic OR)

Suppose F2:F4 contains allowed categories: `"HVAC"`, `"Boiler"`, `"Chiller"`.

```excel
=FILTER(A2:C10, ISNUMBER(MATCH(A2:A10, F2:F4, 0)), "No match")
```

MATCH returns a number for matches → `ISNUMBER` produces a Boolean array.

***

## 1.6 Dynamic Array Behavior

*   The result **expands automatically** (spill behavior).
*   If any cell below the formula is not empty → `#SPILL!`
*   Use spill reference syntax to refer to the full result:

```excel
=SUM(C2#)
```

If `C2` contains a spilling formula, `C2#` means “the whole spill”.

***

# 2. Using `LAMBDA` Inside Array Formulas

`LAMBDA` allows you to create **custom Excel functions**—including fully array-enabled ones.

You can:

*   Use `LAMBDA` inline (anonymous)
*   Name the function using the **Name Manager**
*   Apply it to arrays that spill

***

## 2.1 Basic Inline LAMBDA Example

Compute the square of each value in A2:A10:

```excel
=LAMBDA(x, x^2)(A2:A10)
```

Explanation:

*   The last parentheses contain the “input”
*   `x` becomes the entire array `A2:A10`
*   `x^2` is computed element-wise → returns a spilled array

***

## 2.2 Using LET + LAMBDA in Arrays

To avoid recalculating expressions, combine `LET` with `LAMBDA`.

Example: Normalize a numeric vector:

$$
x_\text{norm} = \frac{x}{\sqrt{\sum(x^2)}}
$$

Formula:

```excel
=LET(
    v, A2:A10,
    norm, SQRT(SUM(v^2)),
    LAMBDA(x, x / norm)(v)
)
```

*   `v` stores the vector
*   `norm` stores its Euclidean norm
*   `LAMBDA(x, x / norm)` scales all components of `v`
*   Result spills

***

## 2.3 Name Manager LAMBDA: reusable custom function

### Step 1 — Create a reusable function

Define it under:  
Formulas → Name Manager → New

**Name:** `NormalizeVector`  
**Refers to:**

```excel
=LAMBDA(v,
    v / SQRT(SUM(v^2))
)
```

### Step 2 — Use it

```excel
=NormalizeVector(A2:A10)
```

Outputs a spilled normalized vector.

***

# 3. Array-Oriented LAMBDA Examples

## 3.1 Dot Product Function

Define:

```excel
=LAMBDA(a, b, SUM(a*b))
```

Usage:

```excel
=DotProduct(A2:A10, B2:B10)
```

***

## 3.2 Apply a Function Row‑wise or Column‑wise (MAP / BYROW / BYCOL)

Dynamic Excel supports:

*   `BYROW(array, lambda)`
*   `BYCOL(array, lambda)`
*   `MAP(array1, array2, ..., lambda)`
*   `REDUCE(initial, array, lambda)`
*   `SCAN(initial, array, lambda)`

### Example: Row sum using BYROW

```excel
=BYROW(A2:C10, LAMBDA(r, SUM(r)))
```

Spills one result per row.

### Example: Apply a transformation to two arrays using MAP

Element‑wise harmonic mean of two vectors:

```excel
=MAP(A2:A10, B2:B10, LAMBDA(x, y, 2/(1/x + 1/y)))
```

***

## 3.3 Build a Flexible Filter Using LAMBDA

Create reusable filter logic:

```excel
=LAMBDA(tbl, col, val,
    FILTER(tbl, col=val)
)
```

Usage:

```excel
=FilterBy(tbl, tbl[Category], "HVAC")
```

***

# 4. Example: Multi-condition Dynamic Filter with LAMBDA and MAP

Assume a criteria table in F2:G4 where:

*   Column F contains column references
*   Column G contains desired values

We build a reusable lambda:

```excel
=LAMBDA(data, cols, vals,
    FILTER(
        data,
        MAP(cols, vals, LAMBDA(c, v, c=v))
    )
)
```

Call it:

```excel
=SmartFilter(A2:C100, F2:F4, G2:G4)
```

This uses:

*   `MAP` to create a Boolean matrix of conditions
*   `FILTER` to extract matching rows
*   Fully dynamic array behavior

***

# If you want, I can also provide:

*   A one‑page cheat sheet of **FILTER**, **MAP**, **BYROW**, **REDUCE**, and **LAMBDA**
*   A working `.xlsx` with examples you can reuse
*   Reusable named Lambdas for regression, normalization, filtering, and matrix ops

Would you prefer a cheat sheet or an Excel workbook?
