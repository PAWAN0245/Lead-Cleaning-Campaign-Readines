# Lead Cleaning & Campaign Readiness (Excel Project)

A small practice project where I clean messy sales-lead data in Excel and find out how many leads are actually ready for a WhatsApp / SMS campaign.

> The data is **sample (synthetic)**: 1,000 leads, Aug to Nov 2026.

## The problem (in simple words)
A company gets leads from Facebook, Google Ads, 99acres and other sources. The data is messy:
- extra spaces and wrong capital letters in names and cities (`mumbai`, `MUMBAI `)
- mobile numbers written in different ways (`+91 98...`, `62213-87837`, `0` at the start)
- invalid numbers (too short, wrong first digit)
- the same number repeated
- some cities are blank

If we send a campaign to this data, many messages fail and money is wasted. So the data must be cleaned first.

## What I did
1. Kept the original data safe in `Raw_Data` (never changed).
2. Cleaned every row with formulas in `Clean_Data`:
   - removed spaces, fixed capital letters (`TRIM`, `PROPER`)
   - kept only the last 10 digits of the mobile (`SUBSTITUTE`, `RIGHT`)
   - checked if the mobile is valid: 10 digits and starts with 6 or more (`IF`, `AND`, `LEN`)
   - marked repeated numbers as Duplicate (`COUNTIF`)
   - blank city became "Unknown"
3. Added the source type (Paid / Portal / Free) from a small table using `VLOOKUP`.
4. Gave every lead one final status: **Invalid Mobile**, **Duplicate** or **Campaign Ready**.
5. Built a `Dashboard` with totals, tables, charts and accuracy checks (`COUNTIF`, `COUNTIFS`).

## Files
| File | What it is |
|---|---|
| `Lead_Cleaning_Basic_Project.xlsx` | The full project |
| `leads_basic_1000.csv` | The raw data (1,000 rows) |

Sheets inside the Excel file:
- **Raw_Data** : original messy data
- **Clean_Data** : formulas that clean each row
- **Source_Master** : small lookup table (source to type)
- **Dashboard** : results and charts
- **How_I_Built_It** : every formula explained
- **Interview_QA** : practice questions and answers

## Results
| Measure | Count |
|---|---|
| Total leads | 1,000 |
| Valid mobiles | 938 |
| Invalid mobiles | 62 |
| Duplicates | 59 |
| **Campaign ready** | **879 (87.9%)** |

Key findings:
- **Facebook** sends the most leads (302) but the worst quality: 81.8% ready, with 38 invalid numbers.
- **Referral** (97.8%) and **Google Ads** (95.3%) give the cleanest data.
- **Mumbai** has the lowest ready rate by city (85.1%); **Pune** is the best of the named cities (92.4%).
- Free sources are cleaner (90.2% ready) than Paid (87.4%) and Portal (87.4%).

## Accuracy checks
The dashboard has 4 checks that must all say **OK**, for example Valid + Invalid = Total, and Ready + Duplicates + Invalid = Total.

## Skills shown
Data cleaning, Excel formulas (TRIM, PROPER, SUBSTITUTE, RIGHT, LEN, IF, AND, IFERROR, COUNTIF, COUNTIFS, VLOOKUP), dashboards, charts, conditional formatting, checking my own work.

## How to explain it to a recruiter (30 seconds)
"I took 1,000 messy sample leads and cleaned them in Excel using formulas, keeping the raw data untouched. I fixed names and cities, standardised mobile numbers, flagged invalid and duplicate numbers, and built a dashboard. It shows that 87.9% of leads are ready for a campaign, and that Facebook leads have the worst data quality. I also added checks so the totals always match."

## How to use
1. Open the Excel file and go to the `Dashboard` sheet.
2. To use new data, paste it into `Raw_Data` (same 6 columns). `Clean_Data` and the dashboard update automatically.
3. If you add more than 1,000 rows, extend the formulas in `Clean_Data` and the ranges in `Dashboard`.
