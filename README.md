# ☕ Coffee Sales Analysis & Dashboard

An end-to-end Excel analysis of 1,000 coffee order line items (Jan 2019 – Aug 2022), covering data cleaning, KPI calculation, and an interactive sales dashboard that shows where revenue comes from and which products and customers drive it.

![Dashboard Preview](images/dashboard_preview.png)

---

## 1. Project Overview

This project analyzes order data from an online coffee retailer selling four coffee types (Arabica, Robusta, Excelsa, Liberica) in three roast levels and four bag sizes to customers in the United States, Ireland, and the United Kingdom.

Raw data was spread across three tables (orders, customers, products). I merged them into a single analysis-ready table, built PivotTables and charts on top of it, and assembled everything into a dashboard.

**Headline numbers**

| Metric | Value |
|---|---|
| Total sales | **$45,134** |
| Orders | 957 |
| Unique customers | 913 |
| Units sold | 3,551 |
| Average order value | $47.16 |
| Period covered | 2 Jan 2019 – 19 Aug 2022 |

---

## 2. Business Problem

Leadership has order-level data but no consolidated view of performance. The questions the business needs answered:

- Which coffee types, roasts, and bag sizes generate the most revenue?
- Which countries are the strongest markets?
- Are sales trending up, and is there seasonality?
- Do loyalty card customers spend more than non-members?
- Who are the most valuable customers, and is revenue concentrated in a few of them?

---

## 3. Project Objectives

1. Combine orders, customers, and products into one clean, analysis-ready dataset.
2. Calculate core KPIs: total sales, orders, customers, units, and average order value.
3. Analyze sales by time, coffee type, roast, size, country, and loyalty status.
4. Identify top customers and measure revenue concentration.
5. Present the results in a single dashboard that non-technical stakeholders can read at a glance.

---

## 4. Dataset Description

Source file: [`data/coffeeOrdersData.xlsx`](data/coffeeOrdersData.xlsx)

| Sheet | Rows | Description |
|---|---|---|
| `orders` | 1,000 | Raw order lines: Order ID, Order Date, Customer ID, Product ID, Quantity |
| `customers` | 1,000 | Customer ID, name, email, phone, address, city, country, postcode, loyalty card (Yes/No) |
| `products` | 48 | Product ID, coffee type, roast, size (kg), unit price, price per 100g, profit per unit |
| `clean data` | 1,000 | Merged, analysis-ready table (16 columns) |
| `TotalSales`, `CountryBySales`, `TopFiveCustomers` | n/a | PivotTables feeding the dashboard |
| `Dashboard` | n/a | Final charts and KPIs |

**Product ID format:** `<Coffee>-<Roast>-<Size>`, e.g. `R-M-1` = Robusta, Medium roast, 1 kg bag.

---

## 5. Tools & Technologies

- **Microsoft Excel**: `XLOOKUP`, `INDEX`/`MATCH`, calculated columns, PivotTables, PivotCharts, slicers
- **Python** (`pandas`, `matplotlib`): used to profile the data, verify the numbers, and generate the static dashboard preview image
- **Git & GitHub**: version control and portfolio hosting

---

## 6. Data Cleaning & Preparation

All steps are visible in the `clean data` sheet as live formulas.

1. **Joined three tables into one.** Customer name, email, and country were pulled from `customers` with `XLOOKUP` on Customer ID. Coffee type, roast, size, and unit price were pulled from `products` with `INDEX`/`MATCH` on Product ID.
2. **Calculated Sales.** `Sales = Unit Price × Quantity` for every line item.
3. **Decoded abbreviations.** Coffee codes (`Ara`, `Rob`, `Lib`, `Exc`) became full names, and roast codes (`L`, `M`, `D`) became Light, Medium, and Dark, so charts and filters are readable.
4. **Handled missing emails.** 206 order lines (20.6%) have no email. The lookup was wrapped in an `IF` so missing values show as blank instead of `0`. Email is not used in the analysis.
5. **Brought in loyalty status.** Loyalty Card (Yes/No) was added from `customers` to enable member vs. non-member comparison.
6. **Validated the data.**
   - 0 fully duplicated rows.
   - 957 unique Order IDs across 1,000 lines. The 43 repeated IDs are multi-item orders, not duplicates, so I kept them.
   - No missing values in any field used for analysis (date, quantity, product, price, sales, country).

---

## 7. Analysis & Key Insights

### Sales trend
| Year | Sales |
|---|---|
| 2019 | $12,187 |
| 2020 | $12,118 |
| 2021 | $13,766 (+13.6% vs. 2020) |
| 2022 (Jan–Aug only) | $7,063 |

- Sales were flat from 2019 to 2020, then grew in 2021.
- 2022 is a partial year, so it should not be compared directly with full years.
- Across all years, **June ($4,843) and March ($4,796)** are the strongest months, and **August ($2,327)** is consistently the weakest.

### Product mix
- **Coffee type:** Excelsa ($12,306), Liberica ($12,054), and Arabica ($11,768) are close together. **Robusta is the laggard at $9,005**, about 27% below Excelsa.
- **Roast:** Light roast leads with **$17,354 (38%)**, ahead of Medium ($14,600) and Dark ($13,179). Order counts are nearly identical across roasts (about 333 each), so Light wins on price per order, not popularity.
- **Size:** **2.5 kg bags generate $23,786, or 53% of all sales**, from roughly the same number of order lines as the 0.2 kg bags, which bring in only 7% ($3,308). Large bags are by far the biggest revenue lever.

### Geography
| Country | Sales | Share |
|---|---|---|
| United States | $35,639 | 79% |
| Ireland | $6,697 | 15% |
| United Kingdom | $2,799 | 6% |

The business is heavily dependent on the US market.

### Customers & loyalty
- Loyalty card holders account for 479 order lines ($20,918); non-members for 521 ($24,216).
- **Members spend slightly less per line item** ($43.67 vs. $46.48), so the loyalty program is not currently lifting basket size.
- The top 5 customers (Allis Wilmore, Brenn Dundredge, Terri Farra, Nealson Cuttler, Don Flintiff) each spent $278–$317, **under 1% of revenue each**. Revenue is spread widely, with little dependence on any single customer.

### Recommendations
1. **Push large bags.** Bundles or discounts on 1 kg and 2.5 kg sizes target the biggest revenue driver.
2. **Run a late-summer campaign.** August is the clear seasonal low.
3. **Grow Ireland and the UK.** Both markets are small relative to the US, so there is room for targeted marketing.
4. **Revisit the loyalty program.** Rewards that encourage larger baskets would give members a reason to spend more per order.
5. **Review Robusta positioning or pricing.** It underperforms the other three types.

---

## 8. Dashboard Preview

![Dashboard Preview](images/dashboard_preview.png)

The Excel dashboard includes:
- KPI cards (total sales, orders, customers, units)
- Monthly sales trend by coffee type
- Sales by country
- Top 5 customers

Open [`data/coffeeOrdersData.xlsx`](data/coffeeOrdersData.xlsx) and go to the **Dashboard** sheet to explore it. The preview image above can be regenerated with `python dashboard_preview.py`.

---

## Repository Structure

```
coffee-sales-analysis/
├── data/
│   └── coffeeOrdersData.xlsx
├── images/
│   └── dashboard_preview.png
├── dashboard_preview.py
└── README.md
```

## How to Use

1. Clone the repo: `git clone https://github.com/oprah-mara/coffee-sales-analysis.git`
2. Open `data/coffeeOrdersData.xlsx` in Excel (365 or 2021+ recommended, since the workbook uses `XLOOKUP`).
3. *(Optional)* Regenerate the preview: `pip install pandas matplotlib openpyxl && python dashboard_preview.py`

## Author

**Amarachi** · Data Science student, CUNY City Tech
GitHub: [@oprah-mara](https://github.com/oprah-mara)
