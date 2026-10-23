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
