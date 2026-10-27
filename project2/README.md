# Project 2: Customer Segmentation and Retention

## Project question

**Which customer groups should Maven Fuzzy Factory focus on to encourage repeat purchases and improve customer retention?**

## Project objective

Use customer order history to group customers, study whether they return to buy again, and prepare a one-page retention recommendation.

## Dataset

- **Dataset:** Toy Store E-Commerce Database (Maven Fuzzy Factory)
- **Source:** [Maven Analytics Data Playground](https://mavenanalytics.io/data-playground/toy-store-e-commerce-database)
- **Files used:** `orders.csv`, `order_items.csv`, `order_item_refunds.csv`, `products.csv`, `website_sessions.csv`, `website_pageviews.csv`, and the data dictionary.

The raw CSV files are stored locally. They are not included in this repository.

## How the data will be used

The main file for customer analysis is `orders.csv`:

- `user_id` identifies a customer.
- `order_id` identifies an order.
- `created_at` gives the order date.
- `price_usd` gives the order value.

Other files provide product, refund, and website-visit details. I will join them only when needed and check that joins do not duplicate orders or inflate totals.

## Planned analysis

### RFM customer segmentation

I will calculate:

- **Recency:** How long it has been since each customer’s most recent order.
- **Frequency:** How many distinct orders each customer placed.
- **Monetary:** The customer’s total order value.

I will use these measures to group customers, then describe which groups are active, frequent, high-value, or need attention.

### Cohort retention

I will group customers by the month of their first order and measure how many place orders in later months.

### Retention plan

After reviewing the results, I will recommend one customer group to focus on, an action to test, and measures to check whether it works.

## Results

Analysis is in progress. I will add checked figures, charts, and findings after completing the analysis.

## Data quality and assumptions

I will check missing IDs and dates, duplicate orders, cancellations or refunds, and unusual order values. I will document the rules used to calculate customer spending and retention.

## Limitations

The dataset represents a fictional online retailer and does not guarantee that the same patterns apply to real or current customers. Recommendations will be based on the available order history and will need testing before being treated as business results.

## AI-use log

I will record how AI tools helped with this project and how I checked any suggested code, calculations, or explanations.












# Project 2: Customer Segmentation and Retention

## Goal
Analyze customer purchasing patterns using RFM segmentation and monthly cohort retention, then propose ways to encourage repeat purchases.

## Dataset
Maven Fuzzy Factory practice dataset. The `orders.csv` file includes `user_id` (customer), `order_id` (order), order date, order value, and cost.

## Work completed
- Calculated each customer's recency, order frequency, and monetary value (RFM).
- Grouped customers into segments.
- Calculated the repeat-customer rate.
- Created a monthly cohort-retention table and heatmap.
- Prepared a retention plan and documented findings and limitations.

## Key results
- Customers analyzed: **31,696**
- Repeat customers: **591**
- Repeat-customer rate: **1.86%**

RFM segment counts:

| Segment | Customers |
|---|---:|
| Lapsed one-time | 15,616 |
| Recent one-time | 8,244 |
| High-value recent one-time | 7,245 |
| High-value active repeat | 348 |
| At-risk repeat | 243 |

## Files
- `project2_customer_retention.ipynb` — analysis notebook
- `findings_and_limitations.md` — findings and caveats
- `retention_plan.md` — proposed small retention test

## Limitation
These are historical results from a practice dataset. The repeat-customer rate covers the full available data period, and order value is not profit.
