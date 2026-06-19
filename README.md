# Part 1: Business Data Cleaning, Validation & Excel Reporting

**Student:** Rohit Kadam
**Student ID:** 2511727
**Repository:** rohitkadam_2511727_part1_data_cleaning
**Tool Used:** Microsoft Excel

---

## Problem Summary

A retail company exported order-level sales data from multiple internal systems. The raw dataset contained several data quality issues including inconsistent text formatting, date format problems, duplicate records, missing values, invalid discounts, calculation mismatches, and order status inconsistencies.

The goal was to clean and validate the dataset, apply business rules, document all issues found, and create summary reports for business review.

---

## Dataset Description

| Property | Details |
|---|---|
| File | raw_orders.xlsx |
| Raw Records | 932 rows |
| Clean Records (after deduplication) | 912 rows |
| Columns | 21 original + 9 calculated |
| Date Range | 2024 – 2025 |
| Categories | Furniture, Office Supplies, Technology |
| Regions | North, South, East, West |

---

## Tools Used

| Tool | Purpose |
|---|---|
| Microsoft Excel | All cleaning, validation, and reporting |
| Excel Functions: TRIM | Remove leading, trailing, and extra spaces |
| Excel Functions: PROPER | Standardize text to Title Case |
| Excel Functions: SUBSTITUTE | Fix double spaces and special characters |
| Excel Functions: IF, OR, AND | Business rule logic and data quality flags |
| Excel Functions: TEXT, YEAR | Extract order month and year |
| Excel Functions: Find & Replace | Standardize inconsistent category names |
| Pivot Tables | Summary reports by region, category, segment, ship mode |

---

## Cleaning Steps Performed

### Task 1 — Raw Data Preserved
- Original file kept as `data/raw_orders.xlsx` — untouched
- All cleaning performed in `data/cleaned_orders.xlsx`

### Task 2 — Text Fields Cleaned
Applied `=TRIM(PROPER(SUBSTITUTE(...)))` formulas directly in columns A–U for 10 text fields:
- `customer_name`, `segment`, `region`, `state`, `city`
- `category`, `sub_category`, `ship_mode`, `payment_status`, `order_status`
- Fixed: extra spaces, leading/trailing spaces, UPPER/lower/mixed case, double spaces in category names

### Task 3 — Dates Cleaned and Validated
- Standardized `order_date` and `ship_date` to consistent `DD/MM/YYYY` format
- Mixed formats found: `21 Jul 2024`, `08/31/2024`, `28-11-2024`, `2024-05-24` — all unified
- Added `shipping_delay_days` column: `=C2-B2`
- Identified 92 records where ship date is before order date — flagged as invalid

### Task 4 — Duplicates Handled
- 20 exact duplicate rows identified and removed
- 24 rows with duplicate `order_id` but conflicting information flagged as `Duplicate - Conflicting`
- No conflicting records silently deleted — all retained with flags

### Task 5 — Business Rules Applied
- Missing `region` (25 rows) → filled as `Unknown`
- Missing `ship_mode` (21 rows) → filled as `Unknown`
- Missing `discount` (26 rows) → filled as `0` (all sales fields valid)
- Negative discount (15 rows) → flagged as `Invalid - Negative Discount`
- Cancelled orders (145) → flagged as `Excluded - Cancelled Order`
- Failed payments (69) → flagged as `Excluded - Failed Payment`
- Refunded orders (71) → flagged as `Separate - Refunded Order`
- Ship before order (92) → flagged as `Invalid - Ship Before Order`

### Task 6 — Calculated Columns Added
| Column | Formula Used |
|---|---|
| `cleaned_discount` | `=IF(P2<0,0,IF(P2>1,1,P2))` |
| `calculated_sales` | `=N2*O2*(1-X2)` |
| `calculated_profit` | `=Y2-R2` |
| `profit_margin` | `=IF(Y2=0,"N/A",Z2/Y2)` |
| `order_month` | `=TEXT(B2,"MMMM")` |
| `order_year` | `=YEAR(B2)` |
| `data_quality_flag` | `=IF(...)` → Clean / Warning / Invalid |

---

## Business Rules Applied

| Rule | Action |
|---|---|
| Missing region | Filled as Unknown |
| Missing ship_mode | Filled as Unknown |
| Missing discount | Treated as 0 where all sales fields valid |
| Negative discount | Flagged as Invalid |
| Cancelled orders | Excluded from completed sales summary |
| Failed payments | Excluded from completed sales summary |
| Refunded orders | Summarized separately |
| Ship before order | Flagged as invalid shipping record |

---

## Summary of Data Quality Issues Found

| Issue | Count |
|---|---|
| Missing region values | 25 |
| Missing ship_mode values | 21 |
| Missing discount values | 26 |
| Negative discount values | 15 |
| Ship date before order date | 92 |
| Exact duplicate rows removed | 20 |
| Conflicting duplicate order IDs | 24 |
| Cancelled orders | 145 |
| Failed payment orders | 69 |
| Refunded orders | 71 |
| Sales/profit calculation mismatches | 56 |
| **Total records after cleaning** | **912** |

---

## Summary of Final Pivot Reports

### 1. Sales and Profit by Region
- South region leads with highest total sales: ₹22,92,699.59
- All regions maintain profit margins between 27–31%

### 2. Sales and Profit by Category and Sub-Category
- Technology is the top category with ₹31,72,752.07 in sales
- Copiers (Technology) is the highest-revenue sub-category

### 3. Order Count by Ship Mode
- Standard Class is the most used ship mode (242 orders)
- Same Day accounts for 204 orders — nearly equal distribution across modes

### 4. Profit Margin by Segment
- Home Office segment has the highest profit margin at 29.90%
- All segments maintain healthy margins between 28–30%

### 5. Refunded/Cancelled/Failed Orders by Region
- North region has the highest number of problem orders (86)
- Cancelled orders account for the largest share of problem records

### 6. Monthly Sales Trend
- May 2024 recorded the highest monthly sales
- Late 2025 months show fewer records due to data cutoff

---

## Key Business Insights

1. **South region drives the most revenue** at ₹22,92,699 — nearly 25% of total sales.
2. **Technology category is the top performer** contributing ₹31,72,752 (35% of total sales).
3. **29.31% overall profit margin** across all valid completed orders.
4. **294 records (32%)** had at least one business rule violation — significant data quality concern.
5. **92 records have ship dates before order dates** — likely a systemic data entry issue in the source system.
6. **Home Office segment** is the most profitable customer segment at 29.90% margin.
7. **56 records show sales/profit mismatches** against calculated values — original values retained for reference.

---

## Assumptions and Limitations

### Assumptions
1. Discount values between 0 and 1.0 (0%–100%) are valid.
2. Missing discount is treated as 0 only when all other sales fields are present and valid.
3. Cancelled and Failed records are retained in dataset but excluded from completed sales summaries.
4. Refunded records are retained and summarised separately.
5. Records with ship date before order date are flagged but not removed.
6. Missing region and ship_mode are non-critical and filled as Unknown.

### Limitations
1. Conflicting duplicate records flagged but not resolved — business team must decide correct version.
2. Ship-before-order records cannot be auto-corrected without knowing which date is erroneous.
3. Sales/profit mismatches retained as-is — original source values kept alongside calculated values.
4. PROPER() function may incorrectly capitalise abbreviations (e.g. "XL" becomes "Xl").

---

## Screenshots Included

| File | Shows |
|---|---|
| `screenshots/raw_data_preview.png` | Raw dataset before cleaning |
| `screenshots/cleaned_data_preview.png` | Cleaned dataset with calculated columns |
| `screenshots/pivot_summary_1.png` | Sales and Profit by Region pivot |
| `screenshots/pivot_summary_2.png` | Monthly Sales Trend pivot |

---

## Repository Structure

```
part1_data_cleaning/
├── data/
│   ├── raw_orders.xlsx
│   └── cleaned_orders.xlsx
├── outputs/
│   ├── data_quality_report.xlsx
│   ├── pivot_summary.xlsx
│   └── cleaning_log.md
├── screenshots/
│   ├── raw_data_preview.png
│   ├── cleaned_data_preview.png
│   ├── pivot_summary_1.png
│   └── pivot_summary_2.png
└── README.md
```
