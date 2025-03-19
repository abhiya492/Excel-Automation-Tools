# Sample Data Files

This directory contains sample data files for demonstrating the Excel Automation Toolkit.

## SalesData.csv

A sample CSV file containing sales data with the following columns:

```
Date,Region,Category,ProductID,Quantity,UnitPrice,Discount,SalesRep
2023-01-05,North,Electronics,E-1001,2,499.99,0.10,John Smith
2023-01-07,South,Electronics,E-1002,1,899.99,0.00,Sarah Johnson
2023-01-10,East,Office Supplies,O-2001,5,24.99,0.05,Michael Brown
2023-01-12,West,Furniture,F-3001,1,799.99,0.15,Emily Davis
2023-01-15,North,Electronics,E-1003,3,299.99,0.00,John Smith
2023-01-18,South,Office Supplies,O-2002,10,12.99,0.00,Sarah Johnson
2023-01-20,East,Electronics,E-1001,1,499.99,0.10,Michael Brown
2023-01-22,West,Furniture,F-3002,2,349.99,0.05,Emily Davis
2023-01-25,North,Office Supplies,O-2003,15,8.99,0.00,John Smith
2023-01-28,South,Electronics,E-1002,2,899.99,0.10,Sarah Johnson
2023-01-31,East,Furniture,F-3001,1,799.99,0.00,Michael Brown
2023-02-03,West,Electronics,E-1003,4,299.99,0.15,Emily Davis
2023-02-05,North,Electronics,E-1001,3,499.99,0.05,John Smith
2023-02-08,South,Office Supplies,O-2001,8,24.99,0.00,Sarah Johnson
2023-02-10,East,Electronics,E-1002,1,899.99,0.10,Michael Brown
```

## EmployeeData.csv

A sample CSV file containing employee information:

```
EmployeeID,FirstName,LastName,Department,HireDate,Salary,Manager
E001,John,Smith,Sales,2019-05-15,65000,Jane Wilson
E002,Sarah,Johnson,Sales,2020-02-20,62000,Jane Wilson
E003,Michael,Brown,Sales,2021-06-10,58000,Jane Wilson
E004,Emily,Davis,Sales,2021-10-05,56000,Jane Wilson
E005,Robert,Wilson,Marketing,2018-03-15,70000,Thomas Lee
E006,Jennifer,Martinez,Marketing,2019-07-22,65000,Thomas Lee
E007,David,Anderson,Marketing,2020-11-18,62000,Thomas Lee
E008,Lisa,Taylor,Finance,2017-09-10,75000,James Miller
E009,James,Thomas,Finance,2018-05-20,72000,James Miller
E010,Mary,Jackson,Finance,2019-08-15,68000,James Miller
E011,William,White,IT,2017-02-28,80000,Patricia Moore
E012,Patricia,Harris,IT,2018-10-12,76000,Patricia Moore
E013,Richard,Clark,IT,2019-04-25,72000,Patricia Moore
E014,Barbara,Lewis,HR,2018-01-15,68000,Linda Young
E015,Joseph,Walker,HR,2019-11-05,64000,Linda Young
```

## InventoryData.csv

A sample CSV file containing inventory information:

```
ProductID,ProductName,Category,SupplierID,UnitPrice,UnitsInStock,ReorderLevel
E-1001,Premium Laptop,Electronics,S001,499.99,25,10
E-1002,Ultra HD TV,Electronics,S001,899.99,15,5
E-1003,Wireless Headphones,Electronics,S002,299.99,50,20
E-1004,Smartphone,Electronics,S003,649.99,30,10
E-1005,Tablet,Electronics,S001,399.99,20,8
O-2001,Desk Organizer,Office Supplies,S004,24.99,100,30
O-2002,Premium Pens (Box),Office Supplies,S005,12.99,200,50
O-2003,Sticky Notes (Pack),Office Supplies,S005,8.99,300,75
O-2004,File Cabinet,Office Supplies,S006,129.99,15,5
O-2005,Paper Shredder,Office Supplies,S004,89.99,20,8
F-3001,Executive Desk,Furniture,S006,799.99,10,3
F-3002,Office Chair,Furniture,S006,349.99,20,8
F-3003,Bookshelf,Furniture,S007,249.99,15,5
F-3004,Conference Table,Furniture,S007,999.99,5,2
F-3005,Filing Cabinet,Furniture,S006,199.99,25,10
```

## ExpenseData.csv

A sample CSV file containing expense information:

```
Date,EmployeeID,Category,Description,Amount,ApprovalStatus
2023-01-10,E001,Travel,Client meeting - airfare,450.00,Approved
2023-01-11,E001,Travel,Client meeting - hotel,225.00,Approved
2023-01-12,E001,Meals,Client dinner,125.50,Approved
2023-01-15,E003,Office Supplies,Printer paper,45.99,Approved
2023-01-18,E007,Marketing,Social media ads,350.00,Approved
2023-01-20,E012,Equipment,Replacement keyboard,85.99,Approved
2023-01-25,E005,Travel,Industry conference - registration,300.00,Approved
2023-01-26,E005,Travel,Industry conference - airfare,375.50,Approved
2023-01-27,E005,Travel,Industry conference - hotel,650.00,Approved
2023-01-30,E008,Subscriptions,Financial analysis software,125.00,Approved
2023-02-05,E002,Travel,Client meeting - train,85.00,Approved
2023-02-05,E002,Meals,Client lunch,78.50,Approved
2023-02-08,E014,Training,HR certification course,450.00,Approved
2023-02-10,E009,Office Supplies,Department supplies,112.43,Approved
2023-02-15,E013,Equipment,External hard drive,129.99,Pending
```

## Using the Sample Data

These sample data files can be used to demonstrate various features of the Excel Automation Toolkit:

1. **Data Processing** - Use the Import CSV function to import these files
2. **Data Cleaning** - Apply data cleaning to fix formatting issues
3. **Reporting** - Generate reports based on the sample data
4. **Validation** - Test data validation rules against these datasets

To use these sample files:

1. Click "Import Data" on the main menu
2. Select the sample CSV file you want to import
3. The data will be imported into a new worksheet
4. You can then use other toolkit functions on the imported data 