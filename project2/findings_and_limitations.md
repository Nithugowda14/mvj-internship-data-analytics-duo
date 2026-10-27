# Project 2: Customer Segmentation and Retention

## Findings

- The dataset contains 31,696 customers.
- 591 customers placed more than one order, giving a repeat-customer rate of 1.86%.
- Customer segments from the RFM analysis:
  - Lapsed one-time: 15,616
  - Recent one-time: 8,244
  - High-value recent one-time: 7,245
  - High-value active repeat: 348
  - At-risk repeat: 243
- Cohort analysis compares the share of each first-order-month group that ordered again in later months. Later-month retention was low in the preview, but the full cohort results should be reviewed before making a final conclusion.

## Recommendations

- Test a small follow-up campaign for one-time customers and compare it with a similar group that does not receive the campaign.
- Contact repeat customers who appear at risk with a relevant reminder or offer.
- Measure repeat orders and revenue after the campaign before expanding it.

## Limitations

- These are historical results from the Maven Fuzzy Factory practice dataset, not current company performance.
- The 1.86% repeat-customer rate counts customers with more than one order during the available dataset period.
- The cohort chart was not completed, so later-month retention findings need verification.
- Order value is not profit. Although the data includes cost fields, this analysis did not calculate profit after refunds and other costs.
- RFM segment labels depend on the scoring rules chosen for this project.
