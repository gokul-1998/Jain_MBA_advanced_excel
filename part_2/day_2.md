# Day 2 — Advanced Excel: Lookups, Conditional Aggregation, Text & Date Functions, Error Handling, PivotTables & PivotCharts

**Batch:** 2nd Year MBA
**Duration:** 3 Hrs
**Format:** Lecture + Hands-on Lab
**Prerequisite:** Day 1 — Excel Basics & Foundation
**Tools (this copy):** LibreOffice Calc 7.x
**Tools (students):** Microsoft Excel 2016+ / 365

> This is your working/trainer copy, written for **LibreOffice Calc** — the app you actually prep and test in. Students are on Microsoft Excel, so hand them `day_2_student_handout.md` / `.pdf`, which uses genuine Excel syntax (XLOOKUP, structured references, built-in PivotTable wizard). Both documents teach the same concepts with the tool each side actually has open — don't mix the mechanics between them.

A hands-on session that takes students from manual nested-IF lookups to proper lookup functions (VLOOKUP, XLOOKUP, INDEX-MATCH), conditional aggregation (SUMIFS/COUNTIFS), text and date manipulation, defensive error handling, and finally PivotTables and PivotCharts — the point where raw data becomes an interactive summary a manager can explore without knowing any formulas.

## Learning Objectives

By the end of this session, students will be able to:

- Use VLOOKUP and XLOOKUP to pull data from a reference table instead of hardcoding values in nested IFs
- Combine INDEX and MATCH for flexible, column-independent lookups
- Aggregate data conditionally with SUMIFS, COUNTIFS, and AVERAGEIFS — answering questions like "total travel spend by the Finance department in Q1"
- Manipulate text with LEFT, RIGHT, MID, LEN, FIND, CONCATENATE/TEXTJOIN, and SUBSTITUTE
- Extract and calculate with dates using YEAR, MONTH, DAY, DATEDIF, EOMONTH, NETWORKDAYS, and TEXT for custom formatting
- Wrap any formula in error-handling functions (IFERROR, IFNA, ISERROR) so a missing lookup or a divide-by-zero never breaks a report
- Build PivotTables from raw transaction data, group by multiple dimensions, add calculated fields, and create PivotCharts that update when the underlying data changes

## Session Run-of-Show

| Time         | Duration | Block                                                | Type     |
| ------------ | -------- | ---------------------------------------------------- | -------- |
| 0:00 – 0:10 | 10 min   | Recap of Day 1 & why lookups replace nested IF       | Concept  |
| 0:10 – 0:35 | 25 min   | VLOOKUP & XLOOKUP                                    | Concept  |
| 0:35 – 0:55 | 20 min   | INDEX-MATCH                                          | Concept  |
| 0:55 – 1:05 | 10 min   | Break                                                | —       |
| 1:05 – 1:30 | 25 min   | SUMIFS, COUNTIFS, AVERAGEIFS                         | Concept  |
| 1:30 – 1:50 | 20 min   | Text & date functions                                | Concept  |
| 1:50 – 2:05 | 15 min   | Error handling: IFERROR, IFNA, ISERROR               | Concept  |
| 2:05 – 2:15 | 10 min   | Break                                                | —       |
| 2:15 – 2:50 | 35 min   | PivotTables & PivotCharts — build, group, chart      | Hands-on |
| 2:50 – 3:00 | 10 min   | Recap, exit ticket & Q&A                             | Wrap-up  |

---

## 1. Recap of Day 1 & Why Lookups Replace Nested IF (0:00 – 0:10)

Open with a quick callback to Day 1's capstone — specifically the nested-IF formula that assigned Regional Targets and Commission Rates:

```
=IF(C2="North",180000,IF(C2="South",150000,IF(C2="East",200000,165000)))
```

Ask: what happens when the company adds a fifth region? A sixth? The formula becomes unreadable, unmaintainable, and error-prone. The solution is a **lookup function** — instead of hardcoding every value inside the formula, store the reference data in a separate table and let the formula *look it up* by key. That's the core idea behind today's session.

**Datasets for today:** `employee_master.csv` (40-row employee reference table with Emp ID, Name, Department, Designation, Hire Date, Base Salary, Bonus %, Region, Manager, Performance Rating, Email, Phone, Status) and `monthly_transactions.csv` (180-row expense transaction log with Txn ID, Emp ID, Transaction Date, Category, Amount, Payment Mode, Approved By, Remarks, Project Code). The transactions reference employees by Emp ID — students will use lookup functions to pull employee details into the transaction sheet, then aggregate and pivot the combined data.

## 2. VLOOKUP & XLOOKUP (0:10 – 0:35)

### VLOOKUP

The classic lookup function, available in every version of Excel (and Calc). It searches the **first column** of a table for a value and returns a value from a specified column in the same row.

```
=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
```

- **lookup_value** — the key to search for (e.g. `E1001`)
- **table_array** — the reference table range, where the **first column** must contain the lookup keys
- **col_index_num** — which column of the table to return (1 = first column, 2 = second, etc.)
- **range_lookup** — `FALSE` for exact match (almost always what you want), `TRUE` for approximate match (sorted data, used for tax brackets / grade bands)

**Worked example on `monthly_transactions.csv`:** The transaction sheet has `Emp ID` in column B but no employee name, department, or salary. Pull the employee's **Last Name** from `employee_master.csv`:

```
=VLOOKUP(B2, employee_master!$A$2:$N$41, 3, FALSE)
```

- `B2` — the Emp ID in the transaction row
- `employee_master!$A$2:$N$41` — the lookup table, absolute-referenced so copying down doesn't shift it
- `3` — column 3 of the lookup table is Last Name
- `FALSE` — exact match

**Calc note:** LibreOffice Calc uses the same `VLOOKUP` syntax. No differences here.

**The three VLOOKUP limitations to teach explicitly:**

1. **Left-only lookup** — the lookup key must be in the *first* column of `table_array`. If you need to look up by Department (column 4), you must rearrange the table or use INDEX-MATCH instead.
2. **Fragile column index** — `col_index_num` is a hardcoded number. Insert a column in the middle of the reference table and every VLOOKUP pointing past that column silently returns the wrong data.
3. **Only returns one value** — if multiple rows match the key, VLOOKUP returns the *first* match only.

**Quick practice (5 min):** In the transactions sheet, add a column for **Department** by writing a VLOOKUP that looks up each Emp ID against the employee master and returns column 4 (Department). Copy it down all 180 rows. Then add another column for **Base Salary** (column 8).

### XLOOKUP (Excel 365 / 2021+)

The modern replacement for VLOOKUP — addresses all three limitations:

```
=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
```

- **lookup_array** — the column to search in (doesn't have to be the first column)
- **return_array** — the column to return from (can be to the left of the lookup column)
- **if_not_found** — a custom message if no match exists (eliminates the need for a separate IFERROR wrapper)
- **match_mode** — 0 for exact (default), -1 for exact or next smaller, 1 for exact or next larger, 2 for wildcard
- **search_mode** — 1 for first-to-last (default), -1 for last-to-first, 2 for binary ascending, -2 for binary descending

**Worked example:** Same task — pull Last Name from the employee master:

```
=XLOOKUP(B2, employee_master!$A$2:$A$41, employee_master!$C$2:$C$41, "Not Found")
```

No column index to break. The lookup column and return column are specified independently. And the `"Not Found"` argument handles missing IDs without a separate error wrapper.

**Calc note:** LibreOffice Calc **does not support XLOOKUP** as of 7.x. Mention this in class, demonstrate it on the projector if running Excel, but always provide the VLOOKUP or INDEX-MATCH equivalent for Calc users. The student handout uses XLOOKUP as the primary syntax since students are on Excel 365.

**Quick practice (5 min):** Rewrite the Department and Base Salary lookups using XLOOKUP. Then write a third XLOOKUP that looks up by **Email** (column 12 in the master) and returns the **Emp ID** — this is a left-direction lookup that VLOOKUP cannot do.

## 3. INDEX-MATCH (0:35 – 0:55)

The workhorse combination that works everywhere, has no column-position dependency, and is the go-to when XLOOKUP isn't available (Calc, Excel 2019 and earlier, Google Sheets pre-2023).

### MATCH

Returns the **position** (row number within a range) of a value:

```
=MATCH(lookup_value, lookup_array, [match_type])
```

- **match_type** — `0` for exact match, `1` for largest value ≤ lookup_value (data must be ascending), `-1` for smallest value ≥ lookup_value (data must be descending)

```
=MATCH("E1010", employee_master!$A$2:$A$41, 0)   → returns 10  (E1010 is the 10th value in the range)
```

### INDEX

Returns the **value** at a given position in a range:

```
=INDEX(return_array, row_num)
```

```
=INDEX(employee_master!$D$2:$D$41, 10)   → returns "IT"  (the 10th value in the Department column)
```

### INDEX-MATCH combined

Feed MATCH's result into INDEX:

```
=INDEX(employee_master!$D$2:$D$41, MATCH(B2, employee_master!$A$2:$A$41, 0))
```

This says: "Find where `B2`'s value sits in the Emp ID column, then return that same row's value from the Department column." No column index, no left-column restriction, no fragility when columns are inserted.

**Worked example:** Pull the **Manager** for each transaction:

```
=INDEX(employee_master!$J$2:$J$41, MATCH(B2, employee_master!$A$2:$A$41, 0))
```

**Two-dimensional INDEX-MATCH** — when you want to look up both the row and the column dynamically (e.g., reading from a rate table where one axis is Region and the other is Rating):

```
=INDEX(rate_table, MATCH(row_key, row_headers, 0), MATCH(col_key, col_headers, 0))
```

This is the formula equivalent of "go to the row labelled *this* and the column labelled *that* and give me the value at the intersection."

**Calc note:** INDEX-MATCH works identically in Calc. This is the recommended lookup method for this trainer copy.

**Quick practice (8 min):** Using INDEX-MATCH, pull the employee's **Hire Date**, **Performance Rating**, and **Status** into three new columns on the transactions sheet. Then write a reverse lookup: given a Manager name in a cell, find the first Emp ID managed by that person.

### When to use which

| Need                                   | Use                       |
| -------------------------------------- | ------------------------- |
| Simple exact lookup, key in first col  | VLOOKUP (FALSE)           |
| Lookup from any column, built-in error | XLOOKUP (Excel 365 only)  |
| Works everywhere, any direction        | INDEX-MATCH               |
| Approximate match (tax brackets, grades) | VLOOKUP (TRUE) or XLOOKUP with match_mode -1 |

---

### Break (0:55 – 1:05, 10 min)

---

## 4. SUMIFS, COUNTIFS, AVERAGEIFS (1:05 – 1:30)

These are the conditional aggregation functions — they answer questions like "what is the total spend on Travel by the Finance department?" without needing to filter the data first.

### SUMIFS

```
=SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)
```

Adds up all values in `sum_range` where **every** criteria pair is satisfied simultaneously.

> **Syntax trap:** `SUMIFS` (with the S, the modern version) puts the **sum range first**, then the criteria pairs. The older `SUMIF` (no S, single condition) puts the **criteria range first**. Teach SUMIFS exclusively — it handles one condition just as well, and students don't need to remember two different argument orders.

**Worked examples on `monthly_transactions.csv`:**

Total spend on **Travel**:
```
=SUMIFS(E2:E181, D2:D181, "Travel")
```

Total spend on **Travel** by employees in the **North** region — this requires pulling Region into the transactions sheet first (via a lookup column added earlier), say it's now in column J:
```
=SUMIFS(E2:E181, D2:D181, "Travel", J2:J181, "North")
```

Total spend in **Q1 2026** (January–March):
```
=SUMIFS(E2:E181, C2:C181, ">="&DATE(2026,1,1), C2:C181, "<="&DATE(2026,3,31))
```

> **The `">"&cell` pattern:** When a criterion involves a comparison operator, write it as a text string concatenated with the value — `">="&DATE(2026,1,1)`. This is not intuitive; walk through it slowly and show what happens if you forget the quotes or the ampersand.

### COUNTIFS

Same syntax, but counts instead of summing — no `count_range` argument, just criteria pairs:

```
=COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2], ...)
```

How many **Credit Card** transactions in **Marketing**:
```
=COUNTIFS(F2:F181, "Credit Card", D2:D181, "Marketing")
```

### AVERAGEIFS

Same syntax as SUMIFS — average range first, then criteria pairs:

```
=AVERAGEIFS(average_range, criteria_range1, criteria1, ...)
```

Average transaction amount for **Hardware** purchases:
```
=AVERAGEIFS(E2:E181, D2:D181, "Hardware")
```

### MAXIFS / MINIFS (Excel 2019+ / 365)

```
=MAXIFS(max_range, criteria_range1, criteria1, ...)
=MINIFS(min_range, criteria_range1, criteria1, ...)
```

Largest single **Training** expense:
```
=MAXIFS(E2:E181, D2:D181, "Training")
```

**Calc note:** LibreOffice Calc supports SUMIFS, COUNTIFS, and AVERAGEIFS. MAXIFS and MINIFS are available from Calc 6.2+ but may need `Ctrl+Shift+Enter` as array formulas in older versions.

**Quick practice (10 min):** Build a small summary table:

| Metric | Formula |
| --- | --- |
| Total Travel spend | `=SUMIFS(...)` on Category = "Travel" |
| Total spend by Finance dept | `=SUMIFS(...)` with the lookup-derived Department column |
| Number of Bank Transfer transactions | `=COUNTIFS(...)` on Payment Mode |
| Average Software License cost | `=AVERAGEIFS(...)` on Category |
| Largest single expense in Q1 | `=MAXIFS(...)` with date criteria |
| Count of transactions by Project P-201 | `=COUNTIFS(...)` on Project Code |

## 5. Text & Date Functions (1:30 – 1:50)

### Text functions

These manipulate the text content of cells — cleaning, extracting, combining.

| Function | Syntax | What it does | Example |
| --- | --- | --- | --- |
| `LEFT` | `=LEFT(text, n)` | First *n* characters | `=LEFT("E1001",1)` → `"E"` |
| `RIGHT` | `=RIGHT(text, n)` | Last *n* characters | `=RIGHT("E1001",4)` → `"1001"` |
| `MID` | `=MID(text, start, n)` | *n* characters from position *start* | `=MID("P-201",3,3)` → `"201"` |
| `LEN` | `=LEN(text)` | Number of characters | `=LEN("Priya Nair")` → `10` |
| `FIND` | `=FIND(find_text, within_text)` | Position of first occurrence (case-sensitive) | `=FIND("@","priya.nair@corp.com")` → `10` |
| `SEARCH` | `=SEARCH(find_text, within_text)` | Same as FIND but case-insensitive | `=SEARCH("NAIR","Priya Nair")` → `7` |
| `SUBSTITUTE` | `=SUBSTITUTE(text, old, new)` | Replaces all occurrences | `=SUBSTITUTE("P-201","P-","PRJ-")` → `"PRJ-201"` |
| `CONCATENATE` | `=CONCATENATE(text1, text2, ...)` | Joins strings | `=CONCATENATE(B2," ",C2)` → `"Priya Nair"` |
| `TEXTJOIN` | `=TEXTJOIN(delimiter, ignore_empty, text1, ...)` | Joins with a delimiter (Excel 2019+) | `=TEXTJOIN(", ",TRUE,A2:A5)` |
| `TEXT` | `=TEXT(value, format)` | Formats a value as text | `=TEXT(42000,"₹#,##0")` → `"₹42,000"` |

**Worked examples on `employee_master.csv`:**

Extract the **numeric part** of an Emp ID:
```
=RIGHT(A2, LEN(A2)-1)     →  "1001"  (everything after the "E")
```

Build a **full name** from First Name and Last Name:
```
=B2&" "&C2                 →  "Priya Nair"
```
Or using TEXTJOIN:
```
=TEXTJOIN(" ", TRUE, B2, C2)
```

Extract the **username** from an email address (everything before `@`):
```
=LEFT(L2, FIND("@", L2)-1)   →  "priya.nair"
```

Extract the **domain** from an email address:
```
=MID(L2, FIND("@", L2)+1, LEN(L2))   →  "corpsolutions.com"
```

**Calc note:** `TEXTJOIN` is available in Calc 6.0+. `CONCATENATE` and `&` work identically. `FIND` is case-sensitive in both; `SEARCH` is case-insensitive in both.

### Date functions

| Function | Syntax | What it returns | Example |
| --- | --- | --- | --- |
| `YEAR` | `=YEAR(date)` | 4-digit year | `=YEAR("12-01-2022")` → `2022` |
| `MONTH` | `=MONTH(date)` | Month number (1–12) | `=MONTH("12-01-2022")` → `1` |
| `DAY` | `=DAY(date)` | Day of month | `=DAY("12-01-2022")` → `12` |
| `TODAY` | `=TODAY()` | Today's date (updates daily) | |
| `DATEDIF` | `=DATEDIF(start, end, unit)` | Difference between dates | `=DATEDIF(F2, TODAY(), "Y")` → years of service |
| `EOMONTH` | `=EOMONTH(start, months)` | End of a month *n* months away | `=EOMONTH(TODAY(), 3)` → end of month 3 months from now |
| `NETWORKDAYS` | `=NETWORKDAYS(start, end)` | Working days between two dates (excl. weekends) | `=NETWORKDAYS("01-01-2026","31-03-2026")` → 64 |
| `TEXT` (dates) | `=TEXT(date, format)` | Custom display | `=TEXT(F2, "MMM YYYY")` → `"Jan 2022"` |

**Worked examples on `employee_master.csv`:**

**Years of service:**
```
=DATEDIF(F2, TODAY(), "Y")
```
This gives the completed years between Hire Date and today. `"M"` gives total months; `"D"` gives total days.

**Month-year label** for a hire date:
```
=TEXT(F2, "MMMM YYYY")    →  "January 2022"
```

**Quarter** from a transaction date:
```
=ROUNDUP(MONTH(C2)/3, 0)   →  1, 2, 3, or 4
```
Or more readably:
```
="Q"&ROUNDUP(MONTH(C2)/3,0)&" "&YEAR(C2)    →  "Q1 2026"
```

**Working days remaining** in the current quarter:
```
=NETWORKDAYS(TODAY(), EOMONTH(TODAY(), 3-MOD(MONTH(TODAY())-1,3)-1))
```

**Calc note:** `DATEDIF` is an undocumented but functional formula in both Excel and Calc — it won't appear in autocomplete, but it works. `NETWORKDAYS` works identically.

**Quick practice (5 min):** On the employee master, add columns for **Years of Service** (`DATEDIF`), **Hire Quarter** (e.g. "Q1 2022"), and **Email Username** (`LEFT` + `FIND`).

## 6. Error Handling: IFERROR, IFNA, ISERROR (1:50 – 2:05)

Errors in Excel aren't bugs — they're signals. But in a dashboard or a report, a raw `#N/A` or `#DIV/0!` looks broken. Error-handling functions let you intercept these and replace them with something meaningful.

### The common errors

| Error | Typical cause |
| --- | --- |
| `#N/A` | VLOOKUP/XLOOKUP can't find the lookup value |
| `#VALUE!` | Wrong data type in a formula (e.g. text where a number is expected) |
| `#REF!` | A referenced cell has been deleted |
| `#DIV/0!` | Division by zero |
| `#NAME?` | Excel doesn't recognise a function name (typo, or function not available) |
| `#NUM!` | Invalid numeric value (e.g. square root of a negative number) |
| `#NULL!` | Incorrect range reference (space instead of colon between cells) |

### IFERROR

The catch-all — intercepts **any** error and returns a fallback value:

```
=IFERROR(formula, value_if_error)
```

```
=IFERROR(VLOOKUP(B2, employee_master!$A$2:$N$41, 4, FALSE), "Unknown Dept")
```

If the lookup fails (employee ID doesn't exist in the master), the cell shows `"Unknown Dept"` instead of `#N/A`.

### IFNA

More precise — catches **only** `#N/A`, lets other errors through so you notice genuine problems:

```
=IFNA(VLOOKUP(B2, employee_master!$A$2:$N$41, 4, FALSE), "ID Not Found")
```

**Best practice:** Use `IFNA` for lookup functions (where `#N/A` is the expected failure mode), and `IFERROR` only when you genuinely want to suppress *all* errors (e.g. a ratio formula where division by zero is a known possibility).

### ISERROR / ISNA (for conditional logic, not suppression)

Returns `TRUE` or `FALSE` — useful inside `IF` when you want to *branch* on an error rather than suppress it:

```
=IF(ISERROR(VLOOKUP(B2, ..., 4, FALSE)), "Needs Review", VLOOKUP(B2, ..., 4, FALSE))
```

> **Note:** This evaluates the VLOOKUP twice. `IFERROR` is more efficient for simple suppression. `ISERROR`/`ISNA` earn their place when you need different logic paths, not just a fallback value.

**A trap to flag explicitly:** wrapping everything in `IFERROR(..., "")` — this hides genuine errors that should be investigated. A formula returning `#REF!` because someone deleted a column should *not* silently become a blank cell. Only suppress errors you've thought through and decided are acceptable.

**Quick practice (5 min):** 
1. Write a VLOOKUP that looks up an Emp ID that doesn't exist in the master — observe the `#N/A`. Wrap it in `IFERROR` with a custom message.
2. Write a formula that divides Total Spend by Number of Transactions — handle the divide-by-zero case when a category has zero transactions.
3. Use `IFNA` on an INDEX-MATCH lookup and verify that a `#VALUE!` error (from deliberately passing text where a number is expected) still shows through.

---

### Break (2:05 – 2:15, 10 min)

---

## 7. PivotTables & PivotCharts (2:15 – 2:50)

This is where everything comes together. A PivotTable takes a flat transaction log and lets you summarise, group, filter, and rearrange it interactively — no formulas, no manual subtotals, no SUMIFS. A PivotChart is a chart bound to a PivotTable that updates automatically when you change the PivotTable's layout.

### What is a PivotTable?

A PivotTable is an interactive summary view of a dataset. You drag fields into four areas:

- **Rows** — the categories down the left side (e.g. Department, Category)
- **Columns** — the categories across the top (e.g. Payment Mode, Quarter)
- **Values** — the numbers being summarised (e.g. Sum of Amount, Count of Txn ID)
- **Filters / Report Filter** — fields used to filter the entire PivotTable without touching the source data (e.g. show only Q1, or only Project P-201)

### Building a PivotTable in Calc

1. Click anywhere inside the dataset (`monthly_transactions.csv`, imported as a range)
2. **Insert → Pivot Table** (Calc calls it Pivot Table, not PivotTable — same thing)
3. In the dialog, drag fields:
   - **Row Fields:** Category
   - **Data Fields:** Amount (defaults to Sum)
   - Click OK
4. Result: a summary showing Total Amount by Category — one row per category, summed automatically

**In Excel (student handout):** Select the data, go to **Insert → PivotTable**, choose "New Worksheet", and drag fields in the PivotTable Field List pane on the right. The field areas are labelled **Rows**, **Columns**, **Values**, and **Filters**.

### Building progressively complex summaries

**Summary 1 — Spend by Category:**
- Rows: Category
- Values: Sum of Amount

**Summary 2 — Spend by Category and Payment Mode (cross-tab):**
- Rows: Category
- Columns: Payment Mode
- Values: Sum of Amount

**Summary 3 — Monthly trend by Category:**
Add a helper column for Month in the source data first: `=TEXT(C2,"MMM-YYYY")` or `=MONTH(C2)`. Then:
- Rows: Category
- Columns: Month
- Values: Sum of Amount

> **Excel 365 date grouping:** In the student handout, students can right-click a date field in the PivotTable, choose **Group**, and group by Months/Quarters/Years directly — no helper column needed. Calc's Pivot Table doesn't support this automatic date grouping, so the helper column is the workaround.

**Summary 4 — Department-level analysis (uses the lookup columns):**
With Department pulled into the transactions sheet via lookup:
- Rows: Department
- Columns: Category
- Values: Sum of Amount
- Filter: Project Code = P-201

**Summary 5 — Count instead of Sum:**
- Rows: Approved By
- Values: Count of Txn ID (change the aggregation from Sum to Count by double-clicking the field in Calc, or clicking the dropdown arrow → Value Field Settings in Excel)

### Calculated fields and value field settings

**Changing aggregation:** Double-click a Value field in the Pivot Table layout (Calc) or click **Value Field Settings** in Excel to change from Sum to Count, Average, Max, Min, etc.

**Show Values As (Excel):** Right-click a value cell → Show Values As → **% of Grand Total**, **% of Row Total**, **% of Column Total**, or **Running Total**. This turns raw numbers into proportions without changing the source data. Calc has limited "Show Values As" support — demonstrate on the projector if running Excel, note the limitation for Calc users.

### Sorting and filtering within PivotTables

- **Sort:** Click any value cell in the PivotTable and sort — it sorts the entire pivot by that column's values
- **Filter:** Use the dropdown arrows on Row/Column labels to filter specific items (e.g. show only "Travel" and "Hardware")
- **Top N filter (Excel):** Value Filters → Top 10 — show only the top 5 categories by spend

### Refreshing a PivotTable

When the source data changes (new rows added, values corrected), the PivotTable does **not** update automatically. You must:
- **Calc:** Right-click the Pivot Table → **Refresh** (or **Data → Refresh**)
- **Excel:** Right-click → **Refresh**, or **PivotTable Analyze → Refresh**

> **Trap to flag:** If new rows are added *below* the original range, the PivotTable won't see them unless the source range is expanded. Converting the source data to a Table (`Ctrl+T` in Excel) before creating the PivotTable avoids this — the Table auto-expands. In Calc, manually update the source range via **Data → Pivot Table → Edit Layout**.

### PivotCharts

A PivotChart is a chart that's tied to a PivotTable — rearranging fields in the PivotTable automatically updates the chart.

**In Excel:**
1. Click inside the PivotTable
2. **PivotTable Analyze → PivotChart** (or **Insert → PivotChart**)
3. Choose a chart type (Column, Bar, Pie, Line)
4. The chart has the same filter buttons as the PivotTable — clicking a filter on the chart filters both chart and table simultaneously

**In Calc:**
1. Select the Pivot Table output range
2. **Insert → Chart** — this creates a regular chart from the pivot output, but it's not dynamically linked to the Pivot Table layout. If you rearrange fields, you'll need to recreate the chart. Mention this limitation explicitly.

**Worked example — build all three together:**

1. Create a PivotTable: Rows = Category, Values = Sum of Amount
2. Insert a PivotChart: Clustered Column chart
3. Add **Payment Mode** to Columns in the PivotTable — watch the chart update to show stacked/grouped columns automatically
4. Use the filter dropdown on the chart to show only "Credit Card" and "Bank Transfer" — watch both table and chart filter simultaneously
5. Sort the PivotTable by Sum of Amount descending — the chart reorders to match

**Quick practice (15 min):**

1. Build a PivotTable showing **Total Spend by Department** (using the lookup-derived column). Which department spends the most?
2. Add **Category** to Columns to create a cross-tab. Which department-category combination has the highest spend?
3. Change the Values aggregation to **Count** — now it shows number of transactions instead of total amount. Which department submits the most expense claims?
4. Add a **Project Code** filter. Filter to Project P-102 only — which department dominates this project's spend?
5. Create a PivotChart (Column chart) from the Department × Category PivotTable. Add a filter for Q1 only (Jan–Mar transactions).
6. Build a second PivotTable: Rows = Approved By, Values = Sum of Amount, Count of Txn ID. Who approves the highest total spend? Who approves the most transactions?

## 8. Recap & Exit Ticket (2:50 – 3:00)

**Five-question exit ticket:**

1. Write a VLOOKUP formula that looks up Emp ID `"E1015"` in the employee master and returns the Department. What would you change to make it return the Base Salary instead?
2. Rewrite the same lookup using INDEX-MATCH. Why is INDEX-MATCH more robust than VLOOKUP when columns might be inserted or deleted?
3. Write a SUMIFS formula that calculates the total amount spent on "Software License" transactions paid by "Bank Transfer".
4. Write a formula using `LEFT` and `FIND` that extracts the username (everything before the `@`) from an email address in cell L2.
5. A VLOOKUP returns `#N/A` for an employee ID that was recently added to the master but hasn't been saved yet. Wrap the VLOOKUP in an appropriate error-handling function. Would you use `IFERROR` or `IFNA` here, and why?

**Looking ahead:** Today's lookup, aggregation, and pivot skills are the analytical core — the next session builds on them with data validation, conditional formatting, what-if analysis, and dashboard design.

---

## 9. Exercises

These exercises use the same two datasets — `employee_master.csv` and `monthly_transactions.csv`. Use them as in-class practice, as reps for faster finishers, or as homework.

### Exercise 1: Lookup Drills

**Files:** `employee_master.csv`, `monthly_transactions.csv`

Using the transactions sheet, add the following columns by looking up each row's Emp ID in the employee master:

| New Column | Lookup Function | What to Return |
| --- | --- | --- |
| Employee Name | VLOOKUP or XLOOKUP | `First Name & " " & Last Name` (combine after lookup) |
| Department | VLOOKUP / XLOOKUP / INDEX-MATCH | Department |
| Designation | INDEX-MATCH | Designation |
| Hire Date | XLOOKUP | Hire Date |
| Base Salary | VLOOKUP | Base Salary |
| Region | INDEX-MATCH | Region |
| Manager | INDEX-MATCH | Manager |
| Performance Rating | XLOOKUP | Performance Rating |

Then answer these questions using the lookup-enriched data:

a) Which employee (by name) has the most transactions?
b) What is the average transaction amount for employees with a Performance Rating of 5?
c) List all transactions by employees whose Status is "On Leave" or "Resigned" — are any of these transactions dated after the employee's last active date?

### Exercise 2: Reverse and Two-Way Lookups

a) **Reverse lookup:** Given a Manager name (e.g. "Kiran Patel"), find the first Emp ID managed by that person. Which function combination works here — VLOOKUP, XLOOKUP, or INDEX-MATCH? Why can't VLOOKUP do this directly?

b) **Two-way lookup:** Create a small summary table with Departments as rows and Categories as columns, and the total spend at each intersection. Use INDEX-MATCH-MATCH to look up the spend for any Department-Category pair from this summary.

c) **Approximate match:** Create a tax-bracket table:

| Income Slab | Tax Rate |
| --- | --- |
| 0 | 0% |
| 30000 | 5% |
| 50000 | 10% |
| 75000 | 15% |
| 100000 | 20% |

Write a VLOOKUP with `TRUE` (approximate match) to find the tax rate for each employee's Base Salary. Then rewrite it using XLOOKUP with `match_mode = -1`.

### Exercise 3: SUMIFS / COUNTIFS / AVERAGEIFS

Build a summary dashboard (a dedicated sheet or section below the data) with these metrics, each computed using a single SUMIFS/COUNTIFS/AVERAGEIFS formula — no filtering, no helper columns:

| # | Metric | Expected Formula Pattern |
| --- | --- | --- |
| 1 | Total spend on Travel | `SUMIFS` — 1 criterion (Category) |
| 2 | Total spend on Marketing by Project P-403 | `SUMIFS` — 2 criteria (Category + Project Code) |
| 3 | Number of Credit Card transactions | `COUNTIFS` — 1 criterion (Payment Mode) |
| 4 | Number of transactions approved by "Fatima Sheikh" in Q1 (Jan–Mar 2026) | `COUNTIFS` — 3 criteria (Approved By + 2 date criteria) |
| 5 | Average amount of Hardware purchases | `AVERAGEIFS` — 1 criterion |
| 6 | Total spend by Finance department in Q2 (Apr–Jun 2026) | `SUMIFS` — 3 criteria (Department + 2 date criteria) |
| 7 | Count of Petty Cash transactions under ₹3,000 | `COUNTIFS` — 2 criteria (Payment Mode + Amount < 3000) |
| 8 | Total Training spend for employees in the North region | `SUMIFS` — 2 criteria (Category + Region) |
| 9 | Number of transactions where Amount exceeds 30,000 | `COUNTIFS` — 1 criterion (Amount > 30000) |
| 10 | Average Software License cost paid by Bank Transfer | `AVERAGEIFS` — 2 criteria (Category + Payment Mode) |

### Exercise 4: Text Functions

Using `employee_master.csv`:

a) **Full Name column:** Combine First Name and Last Name into "Last, First" format (e.g. "Nair, Priya") using `CONCATENATE` or `&` or `TEXTJOIN`.

b) **Extract username from email:** Use `LEFT` and `FIND` to pull everything before the `@` sign.

c) **Extract domain from email:** Use `MID`, `FIND`, and `LEN` to pull everything after the `@` sign.

d) **Mask phone number:** Show only the last 4 digits: `="XXXXXX"&RIGHT(M2,4)` → `"XXXXXX3210"`.

e) **Clean and standardise Designation:** Use `SUBSTITUTE` to replace "Sr." with "Senior" wherever it appears: `=SUBSTITUTE(E2, "Sr.", "Senior")`.

f) **Extract Project Number** from Project Code in the transactions sheet (e.g. `"P-201"` → `"201"`): Use `MID` or `RIGHT` + `LEN`.

g) **Build a sentence:** Create a column that reads: `"[Name] from [Department] spent ₹[Amount] on [Category]"` using concatenation and `TEXT` for number formatting.

### Exercise 5: Date Functions

Using `employee_master.csv` and `monthly_transactions.csv`:

a) **Years of service:** Calculate completed years between Hire Date and today using `DATEDIF(F2, TODAY(), "Y")`.

b) **Months of service:** `DATEDIF(F2, TODAY(), "M")`.

c) **Tenure bucket:** Write an IF that classifies employees as `"< 2 years"`, `"2–5 years"`, or `"5+ years"` based on years of service.

d) **Hire quarter:** `="Q"&ROUNDUP(MONTH(F2)/3,0)&" "&YEAR(F2)` → e.g. `"Q1 2022"`.

e) **Transaction quarter:** Add a Quarter column to the transactions sheet the same way.

f) **Days since last transaction:** For each employee, find their most recent transaction date and calculate the gap to today.

g) **Working days in Q1 2026:** `=NETWORKDAYS(DATE(2026,1,1), DATE(2026,3,31))`.

h) **End of probation:** Assuming a 6-month probation, calculate the probation end date: `=EOMONTH(F2, 6)`.

i) **Formatted date label:** Convert Hire Date to `"15-Jun-2023"` format using `=TEXT(F2, "DD-MMM-YYYY")`.

### Exercise 6: Error Handling

a) In the transactions sheet, change one Emp ID to `"E9999"` (doesn't exist in the master). Run all lookup formulas. Which show `#N/A`? Wrap each in `IFNA` with a descriptive message.

b) Write a formula that calculates **Average Spend per Transaction by Category** using `SUMIFS / COUNTIFS`. For a category with zero transactions, the formula returns `#DIV/0!`. Wrap it in `IFERROR` to return `0`.

c) Create a validation column: `=IF(ISNA(VLOOKUP(B2,...)), "Invalid ID", "OK")` — this flags bad Emp IDs without suppressing the error in the lookup columns.

d) Deliberately nest an `IFERROR` inside another — e.g. "try XLOOKUP first; if that fails, try VLOOKUP on a backup table; if that also fails, return 'Not Found'". Discuss why this is brittle and when it's justified.

### Exercise 7: PivotTable & PivotChart Analysis

**All exercises use `monthly_transactions.csv` enriched with lookup columns from Exercise 1.**

a) **Basic PivotTable:** Rows = Category, Values = Sum of Amount. Sort descending by amount. Which category has the highest total spend?

b) **Cross-tab:** Rows = Department, Columns = Category, Values = Sum of Amount. Which department-category pair dominates the budget?

c) **Time series:** Rows = Month (add a Month helper column or use Excel's date grouping), Columns = Category, Values = Sum of Amount. Is there a spending trend across the 6-month period?

d) **Count vs Sum:** Build a PivotTable with Rows = Approved By, Values = both Sum of Amount AND Count of Txn ID. Who approves the most transactions? Who approves the highest total value? Are these the same person?

e) **Filters:** Add Project Code as a Report Filter. Filter to P-102 — which categories dominate? Switch to P-403 — how does the pattern change?

f) **% of Total (Excel only):** On the Department × Category PivotTable, right-click a value → Show Values As → % of Row Total. Which category represents the largest share of each department's spend?

g) **PivotChart:** From the Category spend PivotTable, insert a **Clustered Column** chart. Add Payment Mode to Columns — the chart should update to show grouped bars. Filter to show only Credit Card and Bank Transfer. Sort by total descending.

h) **Second PivotChart:** Build a PivotTable with Rows = Month, Values = Sum of Amount. Insert a **Line** chart. Add a trendline. Is total monthly spend increasing, decreasing, or stable?

i) **Drill-down (Excel):** Double-click any value cell in a PivotTable — Excel creates a new sheet with just the rows that make up that number. Try it on the highest-spend category cell. How many rows appear?

---

**Trainer prep:** Share `employee_master.csv` and `monthly_transactions.csv` with each student (or pre-load both as tabs in a starter workbook); have a projector with the Ribbon visible at readable zoom. For the PivotTable section, prepare one completed pivot on the projector as a reference — students tend to struggle more with the field-dragging mechanic than with the concept. Keep an answer key for the SUMIFS/COUNTIFS summary and the exit ticket on hand. **This copy is LibreOffice Calc throughout** — if you're demoing live on classroom machines running Microsoft Excel instead, switch to the mechanics in `day_2_student_handout.md` (real XLOOKUP, `Ctrl+T` Tables for auto-expanding pivot source ranges, built-in date grouping in PivotTables, Show Values As) so what's on the projector matches what students see on their own screens.
