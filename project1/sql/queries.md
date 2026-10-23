## SQL Queries Used

The queries below were used to inspect the Online Retail II data and calculate sales and customer metrics. Revenue is calculated as `Quantity × Price`. For sales metrics, cancellation invoices and rows with non-positive quantity or price are excluded.

### 1. Check the invoice date range and row count

```sql
SELECT
    MIN(invoice_date) AS first_date,
    MAX(invoice_date) AS last_date,
    COUNT(*) AS cleaned_rows
FROM transactions_clean;
```

**Purpose:** Checks the earliest and latest invoice dates and how many rows are in the cleaned table.

### 2. Check rows and date range by year

```sql
SELECT
    EXTRACT(YEAR FROM invoice_date) AS year,
    MIN(invoice_date) AS first_date,
    MAX(invoice_date) AS last_date,
    COUNT(*) AS rows
FROM transactions_clean
GROUP BY year
ORDER BY year;
```

**Purpose:** Confirms which years are represented and how many rows belong to each year.

### 3. Check for exact duplicate rows

```sql
SELECT
    COUNT(*) AS total_rows,
    COUNT(DISTINCT (
        invoice, stock_code, description, quantity,
        invoice_date, price_gbp, customer_id, country
    )) AS distinct_rows
FROM transactions_clean;
```

**Purpose:** Compares total rows with distinct rows to check whether exact duplicates remain.

### 4. Count missing values and cancellation invoices

```sql
SELECT
    COUNT(*) AS total_rows,
    COUNT(*) FILTER (
        WHERE customer_id IS NULL
    ) AS missing_customer_id,
    COUNT(*) FILTER (
        WHERE description IS NULL OR TRIM(description) = ''
    ) AS missing_description,
    COUNT(*) FILTER (
        WHERE invoice LIKE 'C%'
    ) AS cancellation_rows
FROM transactions_clean;
```

**Purpose:** Checks for missing customer IDs and descriptions, and counts cancellation rows.

### 5. Check unusual quantity and price values

```sql
SELECT
    COUNT(*) FILTER (WHERE quantity < 0) AS negative_quantity_rows,
    COUNT(*) FILTER (WHERE quantity = 0) AS zero_quantity_rows,
    COUNT(*) FILTER (WHERE price_gbp < 0) AS negative_price_rows,
    COUNT(*) FILTER (WHERE price_gbp = 0) AS zero_price_rows
FROM transactions_clean;
```

**Purpose:** Finds negative or zero quantities and prices for data-quality review.

### 6. Check negative quantities on cancellation invoices

```sql
SELECT
    COUNT(*) FILTER (
        WHERE invoice LIKE 'C%' AND quantity < 0
    ) AS cancellation_with_negative_quantity,
    COUNT(*) FILTER (
        WHERE invoice LIKE 'C%' AND quantity >= 0
    ) AS cancellation_with_nonnegative_quantity,
    COUNT(*) FILTER (
        WHERE invoice NOT LIKE 'C%' AND quantity < 0
    ) AS other_negative_quantity_rows
FROM transactions_clean;
```

**Purpose:** Checks whether negative quantities match cancellation invoices and identifies other negative-quantity rows for review.

### 7. Review negative-price records

```sql
SELECT
    invoice,
    stock_code,
    description,
    quantity,
    price_gbp,
    invoice_date
FROM transactions_clean
WHERE price_gbp < 0
ORDER BY price_gbp;
```

**Purpose:** Shows the records with negative prices so they can be investigated rather than silently changed.

### 8. Count regular sales rows

```sql
SELECT COUNT(*) AS regular_sales_rows
FROM transactions_clean
WHERE invoice NOT LIKE 'C%'
  AND quantity > 0
  AND price_gbp > 0;
```

**Purpose:** Counts rows that meet the rules used for regular sales analysis.

### 9. Calculate sales by month

```sql
SELECT
    STRFTIME(invoice_date, '%Y-%m') AS sales_month,
    ROUND(SUM(quantity * price_gbp), 2) AS sales_gbp
FROM transactions_clean
WHERE invoice NOT LIKE 'C%'
  AND quantity > 0
  AND price_gbp > 0
GROUP BY sales_month
ORDER BY sales_month;
```

**Purpose:** Calculates monthly sales revenue for regular sales rows.

### 10. Calculate month-to-month sales growth

```sql
WITH monthly_sales AS (
    SELECT
        STRFTIME(invoice_date, '%Y-%m') AS sales_month,
        SUM(quantity * price_gbp) AS sales_gbp
    FROM transactions_clean
    WHERE invoice NOT LIKE 'C%'
      AND quantity > 0
      AND price_gbp > 0
    GROUP BY sales_month
),
with_previous_month AS (
    SELECT
        sales_month,
        sales_gbp,
        LAG(sales_gbp) OVER (ORDER BY sales_month) AS previous_month_sales
    FROM monthly_sales
)
SELECT
    sales_month,
    ROUND(sales_gbp, 2) AS sales_gbp,
    ROUND(
        100.0 * (sales_gbp - previous_month_sales)
        / NULLIF(previous_month_sales, 0),
        2
    ) AS mom_growth_percent
FROM with_previous_month
ORDER BY sales_month;
```

**Purpose:** Uses `LAG` to compare each month’s sales with the previous month and calculate percentage growth.

### 11. Calculate repeat-customer percentage

```sql
SELECT
    COUNT(*) AS customers_with_id,
    SUM(CASE WHEN order_count > 1 THEN 1 ELSE 0 END) AS repeat_customers,
    ROUND(
        100.0 * SUM(CASE WHEN order_count > 1 THEN 1 ELSE 0 END)
        / COUNT(*),
        2
    ) AS repeat_customer_percent
FROM (
    SELECT
        customer_id,
        COUNT(DISTINCT invoice) AS order_count
    FROM transactions_clean
    WHERE invoice IS NOT NULL
      AND invoice NOT LIKE 'C%'
      AND customer_id IS NOT NULL  
    GROUP BY customer_id
) AS customer_order_counts;
```

**Purpose:** Counts customers with an ID, identifies customers with more than one distinct non-cancellation invoice, and calculates their percentage.

### Important note

Use the exact column names from your table. If your table uses names such as `"Customer ID"`, `"Invoice"`, or `"Price"` instead of `customer_id`, `invoice`, or `price_gbp`, update the query names to match. The AOV and top-10 product queries were suggested for your next step; include them only after you have run them and checked their results.
