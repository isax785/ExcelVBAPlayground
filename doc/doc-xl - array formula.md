# Array Formula

- [Array Formula](#array-formula)
  - [1) What is an Array Formula?](#1-what-is-an-array-formula)
  - [2) How to Implement \& Run an Array Formula](#2-how-to-implement--run-an-array-formula)
    - [A. Dynamic Arrays (Microsoft 365 / Excel 2021+)](#a-dynamic-arrays-microsoft-365--excel-2021)
    - [B. Legacy CSE Arrays (Excel 2019 and earlier)](#b-legacy-cse-arrays-excel-2019-and-earlier)
  - [3) Basic Array Concepts (with Examples)](#3-basic-array-concepts-with-examples)
  - [4) Common Dynamic Array Functions \& Syntax](#4-common-dynamic-array-functions--syntax)
  - [5) Practical Examples](#5-practical-examples)
    - [5.1 Dot Products (Two Ways)](#51-dot-products-two-ways)
    - [5.2 Linear Regression (Multiple Methods)](#52-linear-regression-multiple-methods)
    - [5.3 Non-Linear Regression (Practical Approaches)](#53-non-linear-regression-practical-approaches)
    - [5.4 Array-Based Filtering \& Aggregation](#54-array-based-filtering--aggregation)
    - [5.5 Building Design Matrices with Dynamic Arrays](#55-building-design-matrices-with-dynamic-arrays)
  - [6) Good Practices \& Pitfalls](#6-good-practices--pitfalls)
  - [7) Quick Reference: Array Syntax Patterns](#7-quick-reference-array-syntax-patterns)
  - [8) Worked Example: Linear Regression with Diagnostics (Dynamic Excel)](#8-worked-example-linear-regression-with-diagnostics-dynamic-excel)
  - [9) Nonlinear Fit via Solver (Template)](#9-nonlinear-fit-via-solver-template)
  - [10) Version Notes](#10-version-notes)
  - [11) Suggested Patterns for Engineering/Analytics](#11-suggested-patterns-for-engineeringanalytics)

---

Below is a concise—but comprehensive—Excel guide to array formulas tailored for modern Excel (Microsoft 365 / Excel 2021+) and backward-compatible with legacy (Ctrl+Shift+Enter, “CSE”) array formulas.

> **Audience:** Power users, analysts, and engineers who want to leverage vectorized, multi-cell calculations (spilled dynamic arrays), and matrix algebra (e.g., dot products, regression).  
> **Excel versions:** Functionality differs between **dynamic array Excel** (Microsoft 365/2021+) and **legacy Excel** (2019 and earlier). I call out both where relevant.

***

## 1) What is an Array Formula?

An **array formula** calculates with multiple values at once—vectors (rows or columns), matrices (2D), or array constants—and can return:

*   a **single result** (e.g., `=SUM(A1:A10*B1:B10)`)
*   a **spilled array** of results (e.g., `=SEQUENCE(10,1)`)

Modern Excel natively supports **dynamic arrays**: many formulas return multiple results that **spill** into neighboring cells automatically. Legacy Excel requires **CSE** for true array evaluation.

***

## 2) How to Implement & Run an Array Formula

### A. Dynamic Arrays (Microsoft 365 / Excel 2021+)

1.  **Enter the formula normally** (press **Enter**). No CSE needed.
2.  If the formula returns multiple results, they **spill** into the adjacent cells.
3.  The original cell is the **anchor**; the spill range is referenced via the **`#`** operator (e.g., `C3#`).
4.  If something blocks the spill (non-empty cell), you’ll see a **`#SPILL!`** error—clear the obstruction or move the formula.
5.  Use **`@`** to enforce implicit intersection (rarely needed; mostly for compatibility).

### B. Legacy CSE Arrays (Excel 2019 and earlier)

1.  Type the formula.
2.  Confirm with **Ctrl+Shift+Enter** (CSE). Excel wraps the formula in `{ }`.
3.  For multi-cell results, **pre-select** the target range, type the formula once, then confirm with CSE.
4.  Edit the whole array range at once; you can’t edit a single cell within a CSE array block.

> **Tip:** Even in modern Excel, CSE still works for backward compatibility, but prefer dynamic arrays.

***

## 3) Basic Array Concepts (with Examples)

*   **Element-wise operations:** `=A1:A10 * B1:B10` creates a 10-row array (one product per row).
*   **Aggregation over arrays:** `=SUM(A1:A10 * B1:B10)` → dot product (single result).
*   **Array constants:** `{1,2,3}` (row), `{1;2;3}` (column), `{1,2;3,4}` (2×2 matrix).  
    Example: `=SUM({1,2,3}*{4,5,6})` returns 32.
*   **Spill references:** If `D2` contains `=SEQUENCE(3,1)`, then `D2#` references the full 3×1 spill.

***

## 4) Common Dynamic Array Functions & Syntax

> *Note:* Use inline code notation inside table cells (not code blocks).

| Category          | Function                       | Syntax (simplified)                                                                              | Purpose                                                   |
| ----------------- | ------------------------------ | ------------------------------------------------------------------------------------------------ | --------------------------------------------------------- |
| Spill / Sequence  | `SEQUENCE`                     | `SEQUENCE(rows, [columns], [start], [step])`                                                     | Generate sequential numbers as an array.                  |
| Spill / Constants | *(n/a)*                        | `{1,2,3}`, `{1;2;3}`, `{1,2;3,4}`                                                                | Inline array constants (row, column, 2D).                 |
| Filter/Sort       | `FILTER`                       | `FILTER(array, include, [if_empty])`                                                             | Filter rows by condition(s).                              |
|                   | `SORT`                         | `SORT(array, [by_index], [order], [by_col])`                                                     | Sort rows or columns.                                     |
|                   | `SORTBY`                       | `SORTBY(array, by_array, [order], ...)`                                                          | Sort by one or more keys.                                 |
| Unique            | `UNIQUE`                       | `UNIQUE(array, [by_col], [exactly_once])`                                                        | Extract distinct or unique-only values.                   |
| Lookup (DA-aware) | `XLOOKUP`                      | `XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])` | Vectorized lookups; spills if `lookup_value` is an array. |
| Reshaping         | `TAKE`                         | `TAKE(array, rows, [columns])`                                                                   | Take top/bottom rows and/or left/right columns.           |
|                   | `DROP`                         | `DROP(array, rows, [columns])`                                                                   | Drop rows/columns.                                        |
|                   | `CHOOSECOLS`                   | `CHOOSECOLS(array, col1, col2, ...)`                                                             | Pick columns by index.                                    |
|                   | `CHOOSEROWS`                   | `CHOOSEROWS(array, row1, row2, ...)`                                                             | Pick rows by index.                                       |
| Reduction         | `SUM`, `AVERAGE`, `MAX`, `MIN` | `SUM(array)`, etc.                                                                               | Aggregate over arrays (often after element-wise ops).     |
| Matrix algebra    | `MMULT`                        | `MMULT(array1, array2)`                                                                          | Matrix multiplication.                                    |
|                   | `MINVERSE`                     | `MINVERSE(array)`                                                                                | Matrix inverse (square non-singular).                     |
|                   | `TRANSPOSE`                    | `TRANSPOSE(array)`                                                                               | Matrix transpose.                                         |
| Statistics        | `LINEST`                       | `LINEST(known_y’s, [known_x’s], [const], [stats])`                                               | Linear regression (returns coefficients; optional stats). |
|                   | `LOGEST`                       | `LOGEST(known_y’s, [known_x’s], [const], [stats])`                                               | Exponential model `y = b*m^x` (linear in log space).      |
| Lambda            | `LAMBDA`                       | `LAMBDA(param1, ..., calculation)`                                                               | Create custom, reusable array-aware functions.            |
| Helpers           | `LET`                          | `LET(name1, value1, ..., calculation)`                                                           | Name arrays/subexpressions for clarity/speed.             |

***

## 5) Practical Examples

### 5.1 Dot Products (Two Ways)

**Data**

*   Vector **a** in `A2:A6`
*   Vector **b** in `B2:B6`

**A) Using `SUMPRODUCT` (robust & fast):**

```excel
=SUMPRODUCT(A2:A6, B2:B6)
```

**B) Using element-wise multiply + `SUM` (dynamic arrays):**

```excel
=SUM(A2:A6 * B2:B6)
```

Both return a scalar (the dot product).

> **Why SUMPRODUCT?** It automatically handles array multiplication + sum, ignores text gracefully, and is typically faster for large ranges.

***

### 5.2 Linear Regression (Multiple Methods)

Assume:

*   Response `y` in `B2:B101`
*   Predictors `X1` in `C2:C101`, `X2` in `D2:D101` (you can have one or many predictors)
*   Optional intercept (constant term)

**Method A: `LINEST` (recommended for quick regression)**

**Coefficients only (slope(s) and intercept):**

```excel
=LINEST(B2:B101, C2:D101, TRUE, FALSE)
```

*   Returns a **row vector**: `[slope_X2, slope_X1, intercept]` if you pass two predictors.
*   The function spills automatically (in dynamic Excel).

**Full stats (coefficients + SE, R², etc.):**

1.  Select a 5-row × (k+1)-column range for output, where `k` = number of predictors.
2.  Enter:
    ```excel
    =LINEST(B2:B101, C2:D101, TRUE, TRUE)
    ```
3.  Press **Enter** (dynamic Excel) or **CSE** (legacy).
4.  Excel returns a statistics table. Use help to interpret rows (coeffs, SE, R², F, df, SSR, SSE).

**Method B: Matrix Algebra (normal equations)**  
This computes the **OLS** solution:

$$
\beta = (X^T X)^{-1} X^T y
$$

You must **prepend a column of ones** to `X` if you want an intercept.

Let’s build `X` with an intercept:

*   In `E2:E101`, set `1` (or use `=1` and fill down).
*   Let `X` be `E2:G101` where columns are `[1, X1, X2]`.

Then enter (in a 3×1 range for 2 predictors + intercept):

```excel
=MMULT(MINVERSE(MMULT(TRANSPOSE(E2:G101), E2:G101)), MMULT(TRANSPOSE(E2:G101), B2:B101))
```

*   Output order: `[intercept; slope_X1; slope_X2]` (rows).

> **Notes:**
>
> *   Requires full column rank (no perfect multicollinearity).
> *   `MINVERSE` is numerically less stable for ill-conditioned matrices; prefer `LINEST` for robustness.

***

### 5.3 Non-Linear Regression (Practical Approaches)

**A) Exponential model `y = b * m^x` with `LOGEST`:**

```excel
=LOGEST(B2:B101, C2:C101, TRUE, TRUE)
```

Interprets `ln(y) = ln(b) + x*ln(m)` internally and returns parameters (and optional stats). With multiple predictors in `X`, use `C2:D101`.

**B) Arbitrary Nonlinear Models → Use Solver or Transformations**
For models like `y = a + b*exp(c*x)` or `y = a / (1 + exp(-b(x-c)))`, closed-form OLS is not available. Use:

1.  A **custom predicted-y column** with the nonlinear formula (using guesses for parameters `a, b, c`).
2.  A **sum of squared residuals** cell: `=SUMXMY2(actual_y_range, predicted_y_range)`.
3.  **Data → Solver**: Minimize SSR by changing parameters (`a, b, c`), subject to constraints if needed.

**C) Linearization where applicable**  
If your model can be linearized (e.g., power law `y = a*x^b` → `ln(y) = ln(a) + b*ln(x)`), transform variables and use `LINEST` on the transformed data.

***

### 5.4 Array-Based Filtering & Aggregation

**Example: Sum revenue for filtered category-month pairs**

*   Data: `Category` in `A2:A1000`, `Month` in `B2:B1000`, `Revenue` in `C2:C1000`.
*   Criteria: `F2` has category, `G2` has month.

```excel
=SUM( (A2:A1000=F2) * (B2:B1000=G2) * C2:C1000 )
```

Dynamic arrays allow direct Boolean arithmetic (TRUE→1, FALSE→0). In older Excel, confirm with **CSE**.

**Same with `SUMPRODUCT` (no CSE):**

```excel
=SUMPRODUCT( (A2:A1000=F2) * (B2:B1000=G2) * C2:C1000 )
```

***

### 5.5 Building Design Matrices with Dynamic Arrays

**Auto-intercept + predictors:**

*   If `X1:Xk` is in `C2:K101`, build `X` with intercept using `HSTACK` (Excel 365 Insider / newer channels) or `CHOOSECOLS` with a helper ones column.

With `HSTACK`:

```excel
=HSTACK( SEQUENCE(ROWS(C2:K101),1,1,0), C2:K101 )
```

This stacks a column of ones with `C2:K101`.

If `HSTACK` isn’t available, place 1’s in a column (say `B2:B101`) and use:

```excel
=CHOOSECOLS(B2:K101, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10)  // or just reference the full range B2:K101
```

***

## 6) Good Practices & Pitfalls

*   **Use `SUMPRODUCT`** for robust array reductions (fewer surprises with text/booleans).
*   **Spill awareness:** Keep spill ranges clear; use `AnchorCell#` in dependent formulas.
*   **Dimensionality:** Ensure conformable shapes for `MMULT` (`m×n` times `n×p`).
*   **Stability:** `MINVERSE` can amplify noise; prefer `LINEST` or `QR`-like stability (only available under the hood).
*   **Types & blanks:** Boolean multiplication requires numeric coercion—structure conditions carefully.
*   **Performance:**
    *   Prefer **structured tables** (`Table1[Column]`) so ranges auto-resize.
    *   Avoid volatile functions where possible.
    *   Reduce repeated calculations with **`LET`**.
*   **Compatibility:** If you must share with legacy Excel users, consider `SUMPRODUCT` instead of element-wise + `SUM`, and avoid functions not available in older versions (e.g., `FILTER`, `UNIQUE`).

***

## 7) Quick Reference: Array Syntax Patterns

| Pattern              | Example                          | Description                            |
| -------------------- | -------------------------------- | -------------------------------------- |
| Element-wise op      | `A2:A10 * B2:B10`                | Pairwise multiply two vectors.         |
| Reduce over array    | `SUM(A2:A10 * B2:B10)`           | Aggregate result to a scalar.          |
| Matrix multiply      | `MMULT(C2:D4, F2:G3)`            | Standard matrix product.               |
| Inverse              | `MINVERSE(C2:E4)`                | Invert 3×3 matrix (non-singular).      |
| Transpose            | `TRANSPOSE(C2:E4)`               | Swap rows/columns.                     |
| Spill ref            | `D2#`                            | Reference entire spill of anchor `D2`. |
| Array constant (row) | `{1,2,3}`                        | 1×3 constant.                          |
| Array constant (col) | `{1;2;3}`                        | 3×1 constant.                          |
| Filter rows          | `FILTER(A2:C1000, A2:A1000="A")` | Return subset by condition.            |

***

## 8) Worked Example: Linear Regression with Diagnostics (Dynamic Excel)

**Goal:** Fit `y` on `X1, X2`, extract coefficients, standard errors, and R².

**Steps:**

1.  **Coefficients & stats** (select a 5×3 block, assuming 2 predictors then enter):
    ```excel
    =LINEST(B2:B101, C2:D101, TRUE, TRUE)
    ```
2.  The first row contains coefficients `[slope_X2, slope_X1, intercept]`.
3.  Use `INDEX` to pick specific statistics from the spill:
    *   Intercept: `=INDEX(LINEST(B2:B101, C2:D101, TRUE, TRUE), 1, 3)`
    *   R² (row 3, col 1 in LINEST layout):  
        `=INDEX(LINEST(B2:B101, C2:D101, TRUE, TRUE), 3, 1)`

> If you frequently need named outputs, wrap with **`LET`** to label the spill and re-use parts cleanly.

***

## 9) Nonlinear Fit via Solver (Template)

1.  **Parameters:** Put guesses in cells, e.g., `a` in `H2`, `b` in `H3`, `c` in `H4`.
2.  **Predicted y:** In `E2`:
    ```excel
    =$H$2 + $H$3 * EXP($H$4 * C2)
    ```
    Fill down to `E101`.
3.  **SSR:** `F2`:
    ```excel
    =SUMXMY2(B2:B101, E2:E101)
    ```
4.  **Solver:** Data → Solver → Set Objective `$F$2` to **Min** by changing `$H$2:$H$4`. Add bounds if needed.

***

## 10) Version Notes

*   **Dynamic arrays** (spilling, `FILTER`, `UNIQUE`, `SORT`, `SEQUENCE`) require **Microsoft 365 / Excel 2021+**.
*   **Legacy Excel** users should use `SUMPRODUCT`, CSE arrays, and helper columns; `LINEST`/`LOGEST` are available but **spill requires CSE** and pre-selecting an output range.

***

## 11) Suggested Patterns for Engineering/Analytics

*   **Dot products / weighted sums:** `SUMPRODUCT(weights, values)`
*   **Vector norms:** `=SQRT(SUM(A2:A101^2))`
*   **Cosine similarity:** `=SUMPRODUCT(a, b) / (SQRT(SUM(a^2))*SQRT(SUM(b^2)))`
*   **Projection of vector `b` onto `a`:** `=SUMPRODUCT(a,b)/SUMPRODUCT(a,a) * a` (returns scalar multiple; spill `*a` as needed)
*   **Normal equation beta:** `=MMULT(MINVERSE(MMULT(TRANSPOSE(X),X)), MMULT(TRANSPOSE(X), y))`
*   **Residuals array:** `=y_range - yhat_range` (spills)
*   **Boolean masks:** `(cond1) * (cond2) * value_range` inside `SUM`/`SUMPRODUCT`

***

If you’d like, I can turn this into:

*   a one-page **cheat sheet PDF**,
*   or a sample **Excel workbook** with the formulas pre-wired for dot products, linear regression (LINEST & matrix), and a nonlinear Solver template.

Would you prefer a PDF cheat sheet, or an `.xlsx` example workbook to explore?
