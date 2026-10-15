# Excel & Functions Fundamentals

> **Veda Technology – Data Analytics Internship · Level 1 · Task 6**
> Practising the everyday Excel formulas used in real analyst work: lookups, logic, conditional aggregation and text cleaning, on a small transactional dataset.

**Author:** Rathod Harshvardhansinh J
**Track:** Data Analytics
**Tools:** Microsoft Excel / Google Sheets

---

## 1. Objective

Build fluency with the formulas an analyst uses every day, and prove it with a workbook that demonstrates **each function on live data**, with short notes on **when to use which formula**.

## 2. What is inside this repository

```
excel-functions-fundamentals/
├── README.md                                  <- you are here
├── workbook/
│   └── Excel_Functions_Fundamentals.xlsx      <- main deliverable (9 sheets, 318 live formulas)
├── screenshots/                               <- one image per sheet (used in the LinkedIn post)
│   ├── 01_overview.png
│   ├── 02_data.png
│   ├── 03_lookups.png
│   ├── 04_logic_if.png
│   ├── 05_conditional_agg.png
│   ├── 06_text_functions.png
│   ├── 07_edge_cases.png
│   └── 08_formula_notes.png

```

## 3. Dataset

The workbook uses a **synthetic, Superstore-style transactional dataset** that I generated for practice (no real customer data):

| Item | Detail |
|---|---|
| Orders | 60 rows, Jul–Sep 2026 |
| Columns | Order ID, Order Date, Customer, Region, Category, Product ID, Qty, Unit Price, Discount %, Sales |
| Regions | West, North, South, East |
| Categories | Furniture, Technology, Office Supplies |
| `Sales` | Formula: `=Qty * Unit Price * (1 - Discount %)` |
| Products table | 10 products with name, category and unit cost (used for lookups) |

You can swap in the public **Sample Superstore** or **Retail Sales** datasets by pasting them over the `Data` sheet – the named ranges (`Region`, `Category`, `Sales`, `Qty`, …) make the formulas easy to repoint.

## 4. Workbook walkthrough

| # | Sheet | What it demonstrates |
|---|---|---|
| 1 | **Overview** | Index of every sheet and a colour legend (blue = input, black = formula, yellow = cells you can change) |
| 2 | **Products** | Lookup table: Product ID → name, category, cost |
| 3 | **Data** | The 60-order dataset with the `Sales` formula and named ranges |
| 4 | **Lookups** | `VLOOKUP` (exact + approximate), `INDEX/MATCH` (including looking *left*), `XLOOKUP` (documented), lookup wrapped in `IFERROR` |
| 5 | **Logic_IF** | `IF`, nested `IF`, `IF + AND`, `IF + OR`, `IFERROR`, and a profit calculation using a lookup |
| 6 | **Conditional_Agg** | `SUMIF`, `SUMIFS`, `COUNTIF`, `COUNTIFS`, `AVERAGEIFS`, date criteria, wildcard criteria, regional summary with a reconciliation check |
| 7 | **Text_Functions** | `LEFT`, `RIGHT`, `MID`, `LEN`, `TRIM`, `UPPER`, `LOWER`, `PROPER`, `SUBSTITUTE`, `FIND`, `&` (concatenation) on messy IDs |
| 8 | **Edge_Cases** | Blank cells, numbers stored as text, trailing spaces breaking lookups, missing keys, divide-by-zero – naive vs robust formula side by side |
| 9 | **Formula_Notes** | When-to-use guide for each formula and answers to the three interview questions |

### 4.1 Lookup functions

| Formula | Example from the workbook | Result |
|---|---|---|
| VLOOKUP (exact) | `=VLOOKUP(B3,ProdTable,2,FALSE)` with `B3 = "P203"` | Mechanical Keyboard |
| INDEX/MATCH | `=INDEX(Products!$B$2:$B$11,MATCH(B3,ProdID,0))` | Mechanical Keyboard |
| Look left | `=INDEX(Products!$A$2:$A$11,MATCH(9,Products!$D$2:$D$11,0))` | P202 (cost → ID, impossible with VLOOKUP) |
| Approximate match | `=VLOOKUP(1820,TierTable,2,TRUE)` | Gold |
| Safe lookup | `=IFERROR(VLOOKUP("P999",ProdTable,2,FALSE),"Not found")` | Not found |
| XLOOKUP | `=XLOOKUP(B3,Products!A2:A11,Products!B2:B11,"Not found")` | Shown as documentation (needs Excel 365 / Google Sheets) |

> The live results in the workbook use `VLOOKUP` and `INDEX/MATCH` so it works in every Excel version and in LibreOffice. The `XLOOKUP` formula is included as text so you can paste it into Excel 365 or Google Sheets.

### 4.2 Logic functions

* `IF(B4>=1000,"High","Normal")` – simple flag.
* `IF(B>=2000,"Large",IF(B>=1000,"Medium",IF(B>=500,"Small","Mini")))` – nested IF with size bands.
* `IF(AND(Region="West",Sales>=800),"Yes","No")` – all conditions must be true.
* `IF(OR(Category="Furniture",Region="North"),"Yes","No")` – any condition is enough.
* `IFERROR(VLOOKUP(...),"Missing")` – friendly message instead of an error.

### 4.3 Conditional aggregation (named ranges)

Named ranges (`Region`, `Category`, `Sales`, `Qty`, `OrderDate`, `Customer`) make the formulas read like sentences:

```excel
=SUMIF(Region,B3,Sales)
=SUMIFS(Sales,Region,B3,Category,B4)
=SUMIFS(Sales,Region,B3,Category,B4,Qty,">="&B5)
=COUNTIFS(Region,B3,Category,B4)
=IFERROR(AVERAGEIFS(Sales,Region,B3,Category,B4),0)
```

Regional summary produced with `SUMIFS` + `COUNTIFS`:

| Region | Total Sales | Orders | Avg Order |
|---|---:|---:|---:|
| West | 7,080.00 | 15 | 472.00 |
| North | 11,270.04 | 15 | 751.34 |
| South | 14,630.09 | 13 | 1,125.39 |
| East | 16,909.81 | 17 | 994.69 |
| **Total** | **49,889.93** | **60** | **831.50** |

A reconciliation cell checks that the four regions add up to `SUM(Sales)` and shows **OK**.

### 4.4 Text functions

Messy IDs such as `"  ord-1001 "` are cleaned with `UPPER(TRIM(...))` → `ORD-1001`. Other examples: `LEFT(…,3)` → `ORD`, `RIGHT(…,4)` → `1001`, `SUBSTITUTE(…,"-","/")` → `ORD/1001`, `FIND("-",…)` → `4`, and `Customer & " – " & Region` to join text.

### 4.5 Edge cases tested

| Scenario | Naive result | Robust approach |
|---|---|---|
| Blank cell inside `IF(I4>0,…)` | Treated as 0, silently wrong | `ISBLANK` check first |
| Number stored as text (`"25"`) in `SUM` | Ignored → 25 instead of 50 | `VALUE()` converts it → 50 |
| `"25" = 25` | FALSE | `VALUE("25") = 25` → TRUE |
| Trailing space in lookup key (`"P203 "`) | Lookup error | `TRIM()` the key first |
| Missing lookup key | Lookup error | `IFERROR` with a clear message |
| Empty group in an average | Divide-by-zero | `IFERROR(…,0)` |
| `COUNTA` vs `COUNT` | Counts text too | Use `COUNT` for numbers only |

## 5. When to use which formula

| Formula | Use it when… | Watch out for |
|---|---|---|
| `VLOOKUP` | The key is in the leftmost column | Use `FALSE` for exact match; hard-coded column index breaks if columns move |
| `XLOOKUP` | You want a modern, any-direction lookup with a built-in "not found" value | Only Excel 365/2021+ and Google Sheets |
| `INDEX/MATCH` | You must look left, or want lookups that survive column insertion; works in all versions | Use `0` in `MATCH` for exact match |
| `IF` | Two outcomes (flag / no flag) | Keep nested IFs to ~3 levels, then use a lookup table |
| `IFERROR` | Replace errors with a friendly value | Do not use it to hide real bugs |
| `SUMIF(S)` | Add values that meet one / many conditions | `SUMIFS` puts `sum_range` **first**; `SUMIF` puts it **last** |
| `COUNTIF(S)` | Count rows meeting conditions | Operators need quotes: `">5"`; join with `&` |
| `AVERAGEIFS` | Average of matching rows | Wrap in `IFERROR` – no matches means divide-by-zero |
| `LEFT/RIGHT/MID` | Pull fixed-position pieces from codes | Combine with `FIND` for variable lengths |
| `TRIM/UPPER/LOWER/PROPER` | Clean imported text before lookups | `TRIM` does not remove non-breaking spaces |

## 6. Interview questions

**1. What's the difference between SUMIF and SUMIFS?**
`SUMIF` takes one condition: `SUMIF(range, criteria, sum_range)`. `SUMIFS` takes many conditions combined with AND: `SUMIFS(sum_range, range1, criteria1, range2, criteria2, …)`. The sum range moves from last to first. Using `SUMIFS` everywhere is a good habit because it also handles a single condition.

**2. How does XLOOKUP improve on VLOOKUP?**
It can look in any direction (not just right), defaults to exact match, has a built-in `if_not_found` argument, uses a lookup column instead of a fragile column number, can return multiple columns, and can search from the bottom up.

**3. When would you use INDEX/MATCH instead of VLOOKUP?**
When the lookup column is to the right of the result column, when columns may be inserted or deleted, when the file must work in older Excel versions without `XLOOKUP`, or for two-way (row + column) lookups.

## 7. Key learnings

* Build and test the lookup first, then wrap it in `IFERROR`, then add `IF` logic – nested IFs stay short and debuggable.
* Most "broken formulas" in real data are really **dirty data**: spaces, text-numbers and blanks. Clean with `TRIM` / `VALUE` before analysing.
* **Named ranges** make formulas self-documenting and easy to audit.
* Always **reconcile**: regional totals must equal the grand total.
* `SUMIFS` + `COUNTIFS` + `AVERAGEIFS` together can replace a small pivot table when you need a formula-driven summary.

## 8. How to use this workbook

1. Download `workbook/Excel_Functions_Fundamentals.xlsx` and open it in Excel or Google Sheets (File → Import).
2. Start on **Overview**, then follow the sheets left to right.
3. Change the **yellow cells** (Product ID, Region, Category, Min Qty, Order value) – all results update automatically.
4. Click any result cell to see the formula behind it.

## 10. Acknowledgements

Created as part of the Veda Technology Data Analytics internship. Dataset is synthetic and generated for learning purposes.
