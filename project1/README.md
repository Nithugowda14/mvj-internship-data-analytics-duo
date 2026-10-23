# Project 1: Retail Sales KPI Analysis

## Question
How do sales change by month and country, and which products sell the most?

## Dataset
Online Retail II from the UCI Machine Learning Repository:
https://archive.ics.uci.edu/dataset/502/online+retail+ii

## Planned work
- Check the data for missing or unusual values.
- Calculate sales value using quantity multiplied by price.
- Summarize revenue, orders, average order value, repeat customers, and top products.
- Make five recommendations based on the results.

## Findings
To be completed after the analysis.

## Limitations
To be completed after checking the data.

## AI use
To be completed. I will record any AI assistance and verify its results.



## Preliminary findings

- Sales value on invoices not marked as cancellations: £20,813,918.43.
- Regular invoices: 45,336.
- Average order value: about £459.10.
- Repeat customers: 4,255 of 5,881 customers with an ID (72.35%).
- Highest monthly sales: November 2011, at £1,509,496.33.
- December 2011 is incomplete, so its monthly sales should not be compared with full months.

## Data quality notes

- Dataset rows: 1,067,371.
- Missing product descriptions: 4,382.
- Missing customer IDs: 243,007. The repeat-customer rate uses only customers with an ID.
- Cancellation invoices: 19,494 rows.
- Negative quantities: 22,950 rows, including 3,457 on invoices not marked as cancellations.
- Zero-price rows: 6,202; negative-price rows: 5.
- No rows have missing invoice numbers or unusable quantity or price values.

These checks are preliminary. Negative quantities and zero prices may have valid explanations, so I have not removed them.




# Project 1: Online Retail II KPI Analysis

An individual internship project analyzing the UCI Online Retail II dataset with Excel, Python/pandas, and SQL. This project examines sales, orders, products, markets, and customer purchasing patterns.

## Business question

**How did sales and customer purchasing patterns change over time, and which products and markets should the retailer focus on to grow sales and improve future profitability?**

**Intended audience:** Retail managers reviewing performance and deciding where to focus attention.

## Files

- [Excel dashboard](dashboard/Online_Retail_Dashboard.xlsx)
- [SQL analysis](sql/retail_analysis.sql)

Open the workbook and go to the **Dashboard** sheet. Use its drop-downs to filter by country, year, and trend metric.

## Key findings

For the dashboard’s default view:

- **Net revenue after returns:** £18,926,477
- **Gross sales:** £19,642,818
- **Orders:** 39,516
- **Average order value:** £497.09, calculated as gross sales divided by orders
- **Customers with an ID:** 5,838
- The United Kingdom contributed approximately **85% of dashboard revenue**.
- The year-over-year comparison uses January through November because December is incomplete in both years.

### Top 10 products by net revenue

| Rank | Stock code | Product | Net revenue | Units sold before returns |
|---:|---|---|---:|---:|
| 1 | 22423 | Regency Cakestand 3 Tier | £314,045 | 26,478 |
| 2 | 85123A | White Hanging Heart T-Light Holder | £248,338 | 94,203 |
| 3 | 85099B | Red Retrospot Jumbo Bag | £178,401 | 96,757 |
| 4 | 47566 | Party Bunting | £147,080 | 28,200 |
| 5 | 84879 | Assorted Colour Bird Ornament | £127,506 | 79,362 |
| 6 | 22086 | Paper Chain Kit 50's Christmas | £116,299 | 35,084 |
| 7 | 79321 | Chilli Lights | £79,905 | 15,841 |
| 8 | 22197 | Small Popcorn Holder | £78,904 | 88,499 |
| 9 | 22386 | Jumbo Bag Pink With White Spots | £75,458 | 39,437 |
| 10 | 84347 | Rotating Silver Angels T-Light Holder | £70,970 | 31,409 |

**Units note:** The dashboard shows net units after returns. The SQL product summary shows units sold before returns, so unit totals can differ. Product revenue shown above is net revenue.

## Recommendation

Run a small, targeted UK sales test using the **Regency Cakestand 3 Tier**, the highest net-revenue product in this analysis at **£314,045**. The UK contributed approximately **85% of dashboard revenue**, so it is a sensible market for a first test.

Before choosing a bundle, use invoice data to find which products customers actually bought with the cakestand. Then test showing those products together to a small group of eligible customers, while keeping a similar group as a comparison. Avoid a broad discount at first.

Compare the groups using **net revenue per customer, units sold, repeat purchases, and returns**. If product cost data is available, also compare **gross profit after the offer**. Expand the offer only if the test improves the result without reducing profit.

### Why this is a test, not a guaranteed profit increase

The dataset shows historical sales and revenue, but it does not include product costs or current demand. It cannot prove which product is most profitable or guarantee that an offer will increase future sales. Check current demand and margins before making a larger investment.

## Dashboard contents

The workbook includes:

- KPI cards for net revenue, orders, average order value, units, active customers, return rate, revenue per customer, and items per order.
- A monthly trend chart with a selectable metric.
- Monthly net revenue comparisons for 2010 and 2011.
- Top international markets and top products.
- Customer segments and top customers.
- A weekday-by-hour purchasing heat map.
- A Notes sheet describing data decisions.

## Data source

- **Dataset:** Online Retail II
- **Publisher:** [UCI Machine Learning Repository](https://archive.ics.uci.edu/dataset/502/online%2Bretail%2Bii)
- **DOI:** [10.24432/C5CG6D](https://doi.org/10.24432/C5CG6D)
- **Download date:** [enter the date you downloaded the dataset]

The raw workbook is not included in this project folder. Download it from the source page above.

## Data preparation

The two year sheets were combined. A stray pivot-table sheet in the workbook was not used.

The preparation and dashboard rules included:

- Removed **22,200 duplicate rows** caused by the overlap between the year sheets on 1–9 December 2010.
- Excluded postage, bank charges, fees, gift vouchers, and zero-price rows, approximately **12,000 rows** in total.
- Accounted for cancellations and returns when calculating net revenue.
- Included transactions without a Customer ID in revenue totals, but excluded them from customer counts and customer segments.
- Compared 2011 with 2010 over January–November because December is incomplete in both years.
- Ranked international markets separately from the United Kingdom because the UK dominates revenue.
- Limited the product ranking to the top 500 products, representing approximately 63% of revenue.
- Used pre-aggregated data in the dashboard workbook rather than storing all raw rows in it.

## Tools and methods

- **Excel:** Dashboard, formulas, and KPI summaries.
- **Python/pandas:** Independent calculations used to check dashboard figures.
- **SQLite:** SQL analysis of average order value and top products, checked against the dashboard.

The SQL script expects a `retail_raw` table with the original columns. Remove the overlapping 1–9 December 2010 rows before loading the data.

## Validation

The workbook was recalculated and reported to have zero formula errors. The default view, Germany in 2011, Bahrain in 2009, rankings, heat map, and year-over-year comparison were checked against separate pandas calculations. The SQL results were also compared with the dashboard figures.

The workbook was visually rendered in LibreOffice. Its appearance has not been independently checked in Microsoft Excel.

## Limitations

- The data is historical and does not explain why customers purchased products.
- Missing Customer IDs mean some revenue cannot be attributed to customers or included in customer segments.
- The dataset represents one retailer and may not reflect other retailers or present-day shopping patterns.
- December is incomplete in both years, so it is excluded from year-over-year comparisons.
- Excluding fees, postage, gift vouchers, and zero-price rows affects the totals.
- The analysis describes observed patterns; it does not prove that one factor caused another.
- Product rankings cover the top 500 products rather than every product.
- Revenue is not profit. Product-cost data is needed to confirm profitability.

## AI-use note

An AI assistant helped explain analytics concepts and draft project documentation. Excel formulas and pandas were used to build and check the dashboard. SQL analysis was later run in SQLite and compared with the dashboard. AI-generated suggestions were reviewed and calculations checked before inclusion.
