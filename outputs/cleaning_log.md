# Cleaning Log — Part 1: Business Data Cleaning, Validation & Excel Reporting

## Assignment Details
- **Student:** Rohit Kadam
- **Student ID:** 2511727
- **Part:** Part 1 — Data Cleaning
- **Tool Used:** Microsoft Excel

---

## Issues Found

| Issue | Count |
|---|---|
| Missing region values | 25 |
| Missing ship_mode values | 21 |
| Missing discount values (filled as 0) | 26 |
| Negative discount values | 15 |
| Ship date before order date | 92 |
| Cancelled orders | 145 |
| Failed payment orders | 69 |
| Refunded orders | 71 |
| Clean records | 618 |
| **Total records** | **912** |

---

## Business Rules Applied

### Rule 1 — Missing Region
- **Action:** All 25 rows with missing region were filled with the value `Unknown`.
- **Flagged in:** `business_rule_flag` column in cleaned_orders.xlsx and quality report.
- **Assumption:** Missing region does not invalidate the record. It is retained but flagged for follow-up.

### Rule 2 — Missing Ship Mode
- **Action:** All 21 rows with missing ship_mode were filled with the value `Unknown`.
- **Flagged in:** `business_rule_flag` column and quality report.
- **Assumption:** Missing ship mode does not invalidate the record. It is retained but flagged.

### Rule 3 — Missing Discount
- **Action:** All 26 rows with missing discount were filled as `0` because all other sales fields (quantity, unit_price, sales, cost, profit) were valid in every case.
- **Assumption:** A missing discount where all other sales fields are valid is treated as no discount applied (0%).

### Rule 4 — Negative Discount
- **Action:** 15 rows with negative discount values were flagged as `Invalid - Negative Discount` in the `business_rule_flag` column.
- **Records removed:** None. Flagged for review only.
- **Assumption:** A negative discount is not a valid business value. It may be a data entry error.

### Rule 5 — Discount Above Allowed Range
- **Action:** No records found with discount above 1.0 (100%). Maximum discount in dataset is 0.65 (65%).
- **Records flagged:** 0.
- **Assumption:** Discount range of 0 to 1.0 (0%–100%) is considered valid.

### Rule 6 — Cancelled Orders
- **Action:** 145 rows with order_status = `Cancelled` were flagged as `Excluded - Cancelled Order`.
- **Records removed:** None. These records are retained in the dataset but excluded from the completed sales summary.
- **Assumption:** Cancelled orders represent lost sales and must not contribute to revenue or profit totals.

### Rule 7 — Failed Payments
- **Action:** 69 rows with payment_status = `Failed` were flagged as `Excluded - Failed Payment`.
- **Records removed:** None. Flagged for exclusion from completed sales summary.
- **Assumption:** Failed payments indicate revenue not actually received.

### Rule 8 — Refunded Orders
- **Action:** 71 rows with payment_status = `Refunded` were flagged as `Separate - Refunded Order`.
- **Records removed:** None. These are retained but must be summarized separately.
- **Assumption:** Refunded orders represent reversed revenue and must be reported independently.

### Rule 9 — Ship Date Before Order Date
- **Action:** 92 rows where ship_date is earlier than order_date were flagged as `Invalid - Ship Before Order`.
- **Records removed:** None. Flagged as invalid shipping records for review.
- **Assumption:** Logically, a shipment cannot happen before an order is placed. These records likely contain data entry errors.

---

## Cleaning Actions Performed

| Task | Action |
|---|---|
| Text fields cleaned | TRIM, SUBSTITUTE, PROPER applied to 10 columns |
| Dates standardised | order_date and ship_date converted to DD/MM/YYYY |
| shipping_delay_days | Calculated as ship_date minus order_date |
| Exact duplicates removed | 20 exact duplicate rows removed |
| Conflicting duplicates flagged | 24 rows flagged as Duplicate - Conflicting |
| Region filled | 25 missing values filled as Unknown |
| Ship mode filled | 21 missing values filled as Unknown |
| Discount filled | 26 missing values filled as 0 |
| Business rule flags | business_rule_flag column added to cleaned_orders.xlsx |

---

## Assumptions Made

1. Discount values between 0 and 1.0 are valid (0% to 100%).
2. Missing discount is treated as 0 only when all other sales fields are present and valid.
3. Cancelled and Failed records are retained in the dataset but excluded from completed sales summaries.
4. Refunded records are retained and summarised separately.
5. Records with ship date before order date are flagged but not removed — they may need manual correction.
6. Missing region and ship_mode are non-critical fields and filled as Unknown rather than invalidating the record.

---

## Records Removed

| Reason | Count |
|---|---|
| Exact duplicate rows | 20 |
| **Total removed** | **20** |

---

## Records Flagged

| Flag | Count |
|---|---|
| Duplicate - Conflicting | 24 |
| Invalid - Negative Discount | 15 |
| Invalid - Ship Before Order | 92 |
| Excluded - Cancelled Order | 145 |
| Excluded - Failed Payment | 69 |
| Separate - Refunded Order | 71 |

---

## Limitations

1. Conflicting duplicate records were flagged but not resolved — business team must decide which version is correct.
2. Ship-before-order records cannot be auto-corrected without knowing which date (order or ship) is erroneous.
3. Discount above allowed range could not be assessed without a formal discount policy document from the business.
4. Text cleaning used PROPER() which may incorrectly capitalise abbreviations (e.g., "XL" becomes "Xl").
