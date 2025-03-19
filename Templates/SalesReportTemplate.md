# Sales Report Template

This Excel template is designed for generating standardized sales reports with automated calculations and formatting.

## Structure

The template includes the following worksheets:

1. **Instructions** - How to use the template
2. **Data Entry** - Where raw sales data is entered or imported
3. **Monthly Summary** - Automatically generated monthly summary
4. **Quarterly Analysis** - Quarterly breakdowns with charts
5. **Dashboard** - Visual overview of key metrics

## Features

### Data Entry Sheet

- Standardized columns with data validation:
  - Date (with date validation)
  - Region (dropdown from predefined list)
  - Product Category (dropdown from predefined list)
  - Product SKU (with format validation)
  - Quantity (numeric validation)
  - Unit Price (currency validation)
  - Discount % (percentage validation)
  - Sales Rep (dropdown from predefined list)

- Automatic calculations:
  - Total Price (Quantity × Unit Price)
  - Discounted Price (Total Price × (1 - Discount%))
  - Profit Margin (based on predetermined costs)

### Monthly Summary Sheet

- Pivot table automatically summarizing:
  - Sales by Region
  - Sales by Product Category
  - Sales by Sales Rep
  - Month-over-month comparison

- Charts:
  - Regional Sales Distribution (pie chart)
  - Product Category Breakdown (bar chart)
  - Daily Sales Trend (line chart)

### Quarterly Analysis Sheet

- Quarter-to-date comparisons
- Quarterly targets vs. actuals
- Top performing:
  - Products
  - Regions
  - Sales Representatives

- Year-over-year quarterly comparison

### Dashboard

- Key performance indicators:
  - Total Sales
  - Average Order Value
  - Items per Order
  - Conversion Rate

- Conditional formatting to highlight:
  - Above target (green)
  - Near target (yellow)
  - Below target (red)

## Usage

1. Enter or import sales data into the Data Entry sheet
2. Use the "Refresh All" button to update calculations and charts
3. Use the "Generate Report" button to create a formatted printable report
4. Use the "Export as PDF" button to save the report as a PDF file

## Customization

The template can be customized by:

1. Modifying the product categories and regions in the "Lists" sheet
2. Adjusting target values in the "Settings" sheet
3. Changing the color scheme in the "Format" section of the "Settings" sheet

## Technical Details

- Uses Excel Tables for data management
- Implements named ranges for formula readability
- Uses structured referencing for calculations
- Incorporates custom VBA functions from the Excel Automation Toolkit:
  - RiskAdjustedValue() for forecast calculations
  - WeightedAverage() for weighted performance metrics
  - FiscalQuarter() for aligned fiscal reporting 