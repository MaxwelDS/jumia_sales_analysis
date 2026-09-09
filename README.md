# Jumia Product Performance Dashboard

### An Excel analysis of price, discount, and rating trends across Jumia Kenya products

## 1. Objective

Online retailers routinely discount products, but a deep discount doesn't automatically translate into customer interest or satisfaction. This project asks:

- Do bigger discounts actually drive more engagement (reviews)?
- Do higher prices correlate with better or worse ratings?
- Which products are highly discounted but poorly rated, or heavily reviewed but under-rated — i.e., which listings need attention?
- What does the overall catalog look like in terms of price, discount, and rating mix?

The goal is a clean, well-documented workbook that a non-technical stakeholder (e.g., a merchandising or category manager) could open and immediately understand.

## 2. Dataset

- **Source:** Product listings scraped from Jumia Kenya (discounted-items category).
- **Raw size:** 115 rows, 6 columns `Product`, `Current price`, `old price`, `Discount`, `Review`, `Ratingd`.
- **Cleaned size:** 112 unique product rows.
- **Fields (raw):**
  | Column | Description |
  |---|---|
  | Product | Product title as listed on Jumia |
  | Current price | Discounted selling price, stored as text (e.g. `KSh 950`) |
  | old price | Original pre-discount price, stored as text |
  | Discount | Advertised discount, as a decimal (e.g. `0.38` = 38%) |
  | Review | Number of customer reviews, some recorded as negative values |
  | Ratingd | Star rating, stored as text (e.g. `4.5 out of 5`) |

Raw data lives in `data/Excel_jumia_dataset.csv`; the cleaned, formula-enriched version lives in the `Cleaned Data` sheet of the workbook.

## 3. Tools

- **Microsoft Excel** data cleaning, formulas, PivotTables, PivotCharts, slicers, dashboard layout.

## 4. Data Cleaning Decisions

Summary:

| Issue                                                       | Rows Affected | Decision                               | Reason                                               |
| ----------------------------------------------------------- | ------------- | -------------------------------------- | ---------------------------------------------------- |
| Negative review counts                                      | 57 rows       | Converted to positive                  | Review counts (a count of people) cannot be negative |
| Duplicate rows (entire row identical)                       | 3 rows        | Removed                                | True duplicates                                      |
| Duplicate product names                                     | 6 rows        | Kept, treated as separate listings     | Same product listed at different prices/variants     |
| One row with a price range instead of a single value        | 1 row         | Replaced with the midpoint             | A single numeric value was required for calculations |
| Extra/irregular spacing in product names                    | Several       | Trimmed                                | Readability and consistent lookups                   |
| Price stored as text with `KSh` prefix and thousands commas | All rows      | Converted to numeric                   | Enable arithmetic and formulas                       |
| Rating stored as text (`"4.5 out of 5"`)                    | All rows      | Numeric rating extracted               | Enable rating calculations and categorisation        |
| Discount stored as decimal                                  | All rows      | Kept as a percentage-formatted decimal | Consistent with Excel's native percentage handling   |

**Missing data:** 58 rows have no `Review` or `Rating` value. These are **not deleted** they are explicitly labelled `"Missing"` by formula in the `Cleaned Data` sheet so they're excluded from rating/engagement analysis without corrupting counts elsewhere.

## 5. Formulas Used

All built in the `Cleaned Data` sheet (Excel table `tblProducts`, `A1:T113`) and the `Analysis` sheet. Key formulas:

**Validation / QA columns**

- `Rating Status` flags ratings outside 0–5 or blank: `=IF(OR(F2<0,F2>5,ISBLANK(F2)),"Check rating","OK")`
- `Discount Status` flags discounts outside 0–100%: `=IF(OR(D2<0,D2>1,ISBLANK(D2)),"Check discount","OK")`
- `Current Price Status` flags cases where the discounted price exceeds the original price: `=IF(B2>C2,"Check prices","OK")`
- `Calculated Discount` independently recomputes discount from the two prices: `=IFERROR((C2-B2)/C2,"")`
- `Discount Check` compares the advertised discount to the calculated one, allowing a 2‑percentage‑point rounding tolerance: `=IF(OR(D2="",K2=""),"Missing",IF(ABS(D2-K2)>2%,"Check Discount","OK"))`

**Categorisation columns**

- `Rating Category` Poor (`<3`), Average (`3–4.5`), Excellent (`>4.5`): `=IF(F2="","Missing",IF(F2<3,"Poor",IF(F2<=4.5,"Average","Excellent")))`
- `Discount Category` Low (`<20%`), Medium (`20–40%`), High (`>40%`): `=IF(D2="","Missing",IF(D2<20%,"Low Discount",IF(D2<=40%,"Medium Discount","High Discount")))`
- `Price Category` Low/Medium/High using dataset quartiles: `=IF(B2="","Missing",IF(B2<=Price_Q1,"Low Price",IF(B2<=Price_Q3,"Medium Price","High Price")))`
- `Engagement Flag flags` products with review counts at or above the 75th percentile: `=IF(E2="","Missing",IF(E2>=Review_Q3,"Strong Engagement","Below Threshold"))`

**Named ranges (defined on the `Analysis` sheet):**

- `Price_Q1` = `QUARTILE.INC(tblProducts[Current Price],1)` = **KSh 493**
- `Price_Q3` = `QUARTILE.INC(tblProducts[Current Price],3)` = **KSh 1,669.50**
- `Review_Q3` = `QUARTILE.INC(tblProducts[Review],3)` = **14 reviews**

**Compound "watch-list" flags** (each combines two conditions to surface products worth a closer look):

- `High Discount + Low Rating`: discount `>40%` AND rating `<3`
- `High Discount + Low Engagement`: discount `>40%` AND reviews below the Q3 threshold
- `Many Reviews + Average Rating`: reviews `≥` Q3 AND rating between 3 and 4.5
- `Strong Engagement + Excellent Rating`: reviews `≥` Q3 AND rating `>4.5`

**Descriptive statistics & correlation (Analysis sheet)** `AVERAGE`, `SUM`, `MAX`/`MIN`, `INDEX`/`MATCH` (to name the max/min-price product), and `CORREL` for the three pairwise relationships analysed (see §7).

## 6. Dashboard Features

The `Dashboard` sheet consolidates the following into a single view (with supporting PivotTable/PivotChart detail on their own sheets):

- **KPI header** total products, average price, average discount, average rating, total reviews, refreshed via `=TODAY()`.
- **Rating Mix** doughnut chart of Poor / Average / Excellent / Missing.
- **Discount Mix** bar chart of Low / Medium / High discount bands.
- **Price vs Rating** average rating by price band.
- **Engagement by Discount** average reviews by discount band.
- **Top 10 by Rating / Top 10 by Reviews / Top 10 by Discount** ranked PivotChart bar charts.
- **Three scatter plots** Discount vs Reviews, Rating vs Reviews, Price vs Rating.
- **Slicers** on Rating Category, Discount Category, and Price Category for interactive filtering across the linked PivotTables.

## 7. Analysis & Key Findings

Based on the cleaned 112-product dataset (54 rows have no review/rating data and are excluded from rating/engagement stats):

- **Catalog snapshot:** average current price **KSh 1,187**, average discount **36.8%**, average rating **3.89**, **723** total reviews across all listed products.
- **Discounts are aggressive:** 55.4% of products (62 of 112) fall into the "High Discount" band (>40% off).
- **Medium discounts drive the most engagement, not the biggest discounts.** Products with 20–40% off average **15.3 reviews**, versus **11.1** for products discounted more than 40% and **9.5** for lightly discounted (<20%) products.
- **Discount depth barely correlates with review count** (Pearson r ≈ **–0.14**, R² ≈ 0.02) bigger discounts do not reliably pull in more reviews.
- **Rating and review volume are essentially uncorrelated** (r ≈ **0.06**), and **price and rating are also weakly related** (r ≈ **0.11**) higher-priced items rate only marginally better (avg. rating 4.08 for High Price vs 3.64 for Low Price), and the relationship is too weak to call a trend.
- **A handful of listings combine high visibility with poor satisfaction** — e.g., a cordless vacuum cleaner has 69 reviews (the most in the dataset) but only a 2.8/5 rating, and several other high-discount items also land in the "Low Rating" watch-list.
- **Top performers exist across categories:** several products (a kettle, a folding hand cart, throw-pillow covers, a storage rack) achieve a perfect 5.0 rating, though typically on a small number of reviews.

The full evidence-to-recommendation trail is documented row-by-row on the `Business insights` sheet.

## 8. Recommendations

1. **Don't rely on discount depth alone to drive engagement** test medium discount bands (20–40%) more deliberately, since they show the strongest reviews-per-listing performance in this dataset.
2. **Audit high-discount, low-rating, high-visibility listings first** (e.g., the 69-review, 2.8-rating vacuum cleaner) investigate quality or listing-accuracy issues, since these products are seen by the most customers.
3. **Compete on perceived value, not price cuts alone** since price shows only a weak positive relationship with rating, better product pages, descriptions, and imagery may matter more than shaving the price further.
4. **Treat "Missing" data as a data-quality signal**, not noise 58 of 115 raw listings had no review/rating data; encouraging reviews on these listings would materially improve analysis coverage.

## 9. Limitations

- **Reviews are a count, not a sales or revenue figure** engagement ≠ conversion, and none of the findings here should be read as profit impact.
- **No product-category field** all products are pooled together; a lamp and a power tool are compared on the same scale.
- **No review text or sentiment** a low rating's cause (defective product, wrong item, shipping issue, etc.) can't be diagnosed from this dataset alone.
- **Snapshot in time** prices, discounts, and review counts reflect a single scrape and will drift.
- **Correlations are weak and the sample is modest** (n = 112, and only 58 rows have both rating and review data) none of the Pearson correlations found are strong enough to imply causation.
- **One row's price required a manual midpoint substitution** (a price range instead of a single value), a small subjective cleaning decision flagged transparently in `Data_Dictionary`.

## 10. File Structure

```
jumia-product-performance-dashboard/
├── README.md
├── data/
│   └── Excel_jumia_dataset.csv       # Raw scraped data
├── dashboard/
│   └── jumia_product_dashboard.xlsx  # Full workbook: raw data, cleaning, pivots, dashboard
└── images/
    ├── raw-data.png                  # Screenshot of the raw dataset
    ├── cleaned-data.png              # Screenshot of the cleaned table with formula columns
    ├── pivot-tables.png              # Screenshot of the supporting PivotTables
    └── dashboard.png                 # Screenshot of the final dashboard
```

## 11. How to Open & Use the Workbook

1. Download `dashboard/jumia_product_dashboard.xlsx` and open it in **Microsoft Excel**.
2. Start on the **Dashboard** sheet for the high-level view; use the slicers to filter by Rating, Discount, or Price category.
3. To see how a figure was calculated, click into the relevant cell — every KPI, category, and flag is a live formula, not a hardcoded value.
4. The **Cleaned Data** sheet (`tblProducts`) is the single source of truth all PivotTables and charts pull from; the **Excel_jumia_dataset** sheet preserves the original, unmodified raw data for reference.
5. See **Data_Dictionary** for the full list of cleaning decisions and threshold definitions, and **Business insights** for the findings-to-recommendations table.
