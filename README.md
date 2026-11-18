# Lead Cleaning & Campaign Readiness Dashboard

An Excel project that cleans messy sales-lead data and shows how many leads are ready for a WhatsApp / SMS campaign.

> **Note:** All data is sample (synthetic) and made for practice. It has 1,000 leads from Aug to Nov 2026.

## Table of Contents
- [Problem](#problem)
- [What This Project Does](#what-this-project-does)
- [Project Files](#project-files)
- [Excel Workbook Structure](#excel-workbook-structure)
- [Cleaning Rules](#cleaning-rules)
- [Results](#results)
- [Excel Functions Used](#excel-functions-used)
- [How to Use](#how-to-use)

## Problem
A company collects leads from Facebook, Google Ads, 99acres, MagicBricks, its website and referrals. The data is messy:

- Extra spaces and inconsistent capital letters in names and cities (`mumbai`, `MUMBAI `)
- Mobile numbers in different formats (`+91 98...`, `62213-87837`, leading `0` or `91`)
- Invalid mobile numbers
- Repeated mobile numbers
- Blank cities

Sending a campaign to this data wastes money because many messages fail. The data must be cleaned first.

## What This Project Does
1. Keeps the original data untouched in a `Raw_Data` sheet.
2. Cleans every row using formulas in a `Clean_Data` sheet.
3. Adds the source type (Paid / Portal / Free) with a lookup table.
4. Gives each lead one final status: **Campaign Ready**, **Invalid Mobile** or **Duplicate**.
5. Shows the results on a `Dashboard` with KPI cards, tables, charts and accuracy checks.

## Project Files
```
.
├── Lead_Cleaning_Basic_Project.xlsx   # Full project (raw data, cleaning, dashboard)
├── leads_basic_1000.csv               # Raw sample data (1,000 rows)
└── README.md
```

**Raw data columns:** `Lead ID`, `Name`, `Mobile`, `City`, `Lead Source`, `Lead Date`

## Excel Workbook Structure
| Sheet | Purpose |
|---|---|
| `Raw_Data` | Original messy data. Never edited. |
| `Clean_Data` | Formulas that clean each row (row 2 matches row 2 of `Raw_Data`). |
| `Source_Master` | Lookup table: lead source to type (Paid / Portal / Free). |
| `Dashboard` | KPIs, city table, source quality table, source type table, charts, checks. |
| `How_I_Built_It` | Every formula and range explained step by step. |
| `Interview_QA` | Practice questions and answers about the project. |

## Cleaning Rules
| Field | Rule | Formula idea |
|---|---|---|
| Name | Remove extra spaces, proper case | `PROPER(TRIM(...))` |
| Mobile | Remove spaces, `-` and `+`, keep the last 10 digits | `RIGHT(SUBSTITUTE(...),10)` |
| Mobile status | Valid if 10 digits and first digit is 6 or higher | `IF(AND(LEN(...)=10, VALUE(LEFT(...,1))>=6), "Valid", "Invalid")` |
| City | Trim, proper case, blank becomes `Unknown` | `IF(TRIM(...)="","Unknown",PROPER(TRIM(...)))` |
| Lead source | Trim extra spaces | `TRIM(...)` |
| Source type | Look up in `Source_Master` | `VLOOKUP(..., FALSE)` |
| Duplicate flag | Second and later appearance of a mobile is a duplicate | `COUNTIF($C$2:C2, C2) > 1` |
| Final status | Invalid first, then Duplicate, else Campaign Ready | Nested `IF` |

Each lead gets exactly one final status, so the totals always add up.

## Results
| Measure | Count |
|---|---|
| Total leads | 1,000 |
| Valid mobiles | 938 |
| Invalid mobiles | 62 |
| Duplicates | 59 |
| **Campaign ready** | **879 (87.9%)** |

**Lead source quality**

| Source | Total | Invalid | Duplicates | Ready | Ready % |
|---|---|---|---|---|---|
| Facebook | 302 | 38 | 17 | 247 | 81.8% |
| Google Ads | 212 | 10 | 0 | 202 | 95.3% |
| 99acres | 151 | 5 | 12 | 134 | 88.7% |
| MagicBricks | 142 | 6 | 14 | 122 | 85.9% |
| Website | 104 | 1 | 16 | 87 | 83.7% |
| Referral | 89 | 2 | 0 | 87 | 97.8% |

**Key findings**
- Facebook sends the most leads but has the lowest ready rate (81.8%), mainly because of invalid numbers.
- Referral (97.8%) and Google Ads (95.3%) give the cleanest data.
- Website leads have the most duplicates relative to their size (16 of 104).
- By city, Mumbai has the lowest ready rate (85.1%) and Pune the highest among named cities (92.4%).

## Accuracy Checks
The dashboard has four checks that should all show **OK**:
- Valid + Invalid = Total leads
- Ready + Duplicates + Invalid = Total leads
- City table total = Total leads
- Source table total = Total leads

## Excel Functions Used
`TRIM`, `PROPER`, `SUBSTITUTE`, `RIGHT`, `LEFT`, `LEN`, `VALUE`, `IF`, `AND`, `IFERROR`, `COUNTIF`, `COUNTIFS`, `COUNTA`, `SUM`, `VLOOKUP`, conditional formatting, filters and charts.

## How to Use
1. Download `Lead_Cleaning_Basic_Project.xlsx` and open it in Excel.
2. Go to the `Dashboard` sheet to see the results.
3. To test with new data, paste it into `Raw_Data` using the same 6 columns. `Clean_Data` and the `Dashboard` update automatically.
4. For more than 1,000 rows, copy the `Clean_Data` formulas further down and extend the ranges in `Dashboard`.

## Limitations
- Practice project on synthetic data, not real customer data.
- Duplicates are detected by mobile number only.
- Formulas are set for 1,000 rows (rows 2 to 1001).

## Author
**Pawan Gupta**
