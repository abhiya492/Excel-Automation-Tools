# Excel Automation Toolkit - User Guide

## Introduction

The Excel Automation Toolkit is a comprehensive set of Excel VBA tools designed to automate data processing tasks, generate reports, and implement data validation. This guide will help you understand how to use each component of the toolkit effectively.

## Getting Started

1. Open the `ExcelAutomationToolkit.xlsm` file
2. Click "Enable Macros" when prompted
3. The main menu will appear automatically. If not, press Alt+F8 and run the `CreateMainMenu` macro

## Main Features

### Data Processing

#### Import Data
- Imports data from CSV files into Excel
- Automatically creates a new worksheet named with the current date
- Auto-fits columns for better readability

**Usage:**
1. Click "Import Data" on the main menu
2. Browse to select the CSV file
3. The data will be imported into a new worksheet

#### Clean Data
- Removes duplicate records
- Trims whitespace from text cells
- Standardizes text formats

**Usage:**
1. Click "Clean Data" on the main menu
2. Select the range of data to clean
3. The function will clean the data and report completion

#### Batch Process Files
- Processes multiple Excel files from a folder
- Creates a summary of all processed files
- Reports on success or failure of each file

**Usage:**
1. Click "Batch Process Files" on the main menu
2. Select a folder containing Excel files
3. A summary sheet will be created with results

### Reporting

#### Generate Monthly Report
- Creates a summary report using pivot tables
- Includes charts for visual representation
- Formats data for readability

**Usage:**
1. Ensure you have a sheet named "Data" with appropriate headers
2. Click "Generate Report" on the main menu
3. The report will be generated in a new sheet named "Monthly_Report"

#### Export as PDF
- Exports the current worksheet as a PDF file
- Preserves formatting and layout
- Opens the PDF automatically after export

**Usage:**
1. Select the worksheet you want to export
2. Click "Export as PDF" on the main menu
3. Choose a location to save the PDF file

#### Schedule Reports
- Provides instructions for setting up automated report generation
- Uses Windows Task Scheduler for automation
- Configurable timing options

**Usage:**
1. Click "Schedule Reports" on the main menu
2. Follow the instructions in the dialog box to set up a scheduled task

### Data Validation

#### Validate Data
- Checks data against predefined validation rules
- Highlights cells with errors
- Creates a validation report with details

**Usage:**
1. Ensure your data has appropriate headers
2. Click "Validate Data" on the main menu
3. Review the validation report in the "ValidationErrors" sheet

#### Apply Validation Rules
- Adds Excel data validation to selected cells
- Prevents entry of invalid data
- Provides user feedback for data entry

**Usage:**
1. Click "Apply Validation Rules" on the main menu
2. Select the range to apply validation to
3. Choose the type of validation to apply
4. Configure validation options as needed

## Data Entry Form

The toolkit includes a data entry form for structured data input:

1. Click Alt+F8 and run the `CreateDataEntryForm` macro
2. Fill in the required fields
3. Click "Submit" to add the data to the Data sheet
4. Use "Clear" to reset the form or "Close" to return to the main menu

## Custom Functions

The toolkit also includes several custom Excel functions that can be used in your worksheets:

- `RiskAdjustedValue(value, riskFactor)` - Calculates risk-adjusted values
- `WeightedAverage(valuesRange, weightsRange)` - Calculates weighted averages
- `FiscalQuarter(date, startMonth)` - Determines fiscal quarter for a date
- `BusinessDaysDifference(startDate, endDate)` - Counts business days between dates
- `FormatCurrency(value, symbol, decimals)` - Standardized currency formatting

## Troubleshooting

### Common Issues

1. **Macros Disabled**
   - Ensure macros are enabled in Excel
   - Set macro security to "Enable all macros" or "Disable with notification"

2. **Missing References**
   - If you get a "Missing reference" error, go to Tools > References and check:
     - Microsoft Office Object Library
     - Microsoft Forms 2.0 Object Library

3. **Data Format Issues**
   - Ensure your data has the expected column headers
   - Check that date fields are properly formatted as dates

### Getting Help

If you encounter issues not covered in this guide, press the "Help" button on the main menu for additional information.

## Best Practices

1. **Back Up Your Data**
   - Always make backups before running batch operations

2. **Test on Sample Data**
   - Test functions on a small dataset before applying to large datasets

3. **Customize Validation Rules**
   - Adjust validation rules in the DataValidation module to match your specific requirements

4. **Create Templates**
   - Save commonly used report configurations as templates

## Advanced Customization

Advanced users can modify the VBA code to customize functionality:

1. Press Alt+F11 to open the VBA editor
2. Navigate to the module corresponding to the function you want to modify
3. Make your changes and save the workbook

## Conclusion

The Excel Automation Toolkit streamlines common Excel tasks, reduces manual work, and ensures data consistency. By following this guide, you should be able to leverage all of its powerful features to improve your productivity. 