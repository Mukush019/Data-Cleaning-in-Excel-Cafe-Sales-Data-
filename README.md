# Data Cleaning and Visualization of Cafe Sales Data

### Project Overview

Real-world datasets are rarely clean, and messy data is one of the common challenges data analysts face before drawing any meaningful insights. In this project, I work with a "Dirty Cafe Sales" dataset, a sales record containing common real-world data quality issues like missing values, inconsistency in formatting, and incorrect data types.

The main objective of this project is to clean and prepare the dataset while minimizing data loss(data deletion). Normally, blank columns are often dropped, a mistake that can be avoided in data cleaning, and one of the goals here is to demonstrate techniques that preserve as much data as possible.

Two common imputation methods, mean imputation and median imputation, are going to be the key focus of this project. By imputing blank values using these methods, I am going to evaluate how each method affects the overall distribution, summary statistics, and visual outcomes of the sales data. This comparison aims to highlight the practical differences between the two approaches and build a stronger understanding of which method is more appropriate. 

The cleaned data will then be explored through EDA to uncover patterns and trends in the cafe sales, such as the best-selling item, sales trends in the year 2023, and revenue distribution. These reports will then be brought together into an interactive Excel dashboard, allowing stakeholders to visually explore cafe sales performance at a glance.

This project in it entirity, from data cleaning to the interactive dashboard, will be executed exclusively in **Microsoft Excel**, showcasing not only how far Excel built-in tools can be pushed to handle a full structure data analysis workflow, but also my expertise in Excel formulas, PivotTables, PivotCharts, conditional formatting, and data validation.

**Objectives**
- Clean the dirty sales dataset while minimizing the amount of data deleted.
- Compare the effects of mean imputation vs median imputation on missing values.
- Perform EDA to explore trends and patterns.
- Build an interactive dashboard

### Data Source
The data used in this project is the **Dirty Cafe Sales** dataset from Kaggle. The data is *synthetic*, intentionally designed to simulate real-world data issues, making it ideal for practicing data cleaning and preparation techniques.

#### Data Summary
- Source: Kaggle
- Type: Synthetic data
- Format: XLS
- Size: 10000 rows by 8 columns
- Time: This data represents one year of sales(2023)

##### Columns in the dataset
* `Transaction ID` - Unique identifier for the sales made at the cafe.
* `Item` - what a customer bought at the cafe.
           That includes: `Coffee`, `Tea`, `Salad`, `Sandwich`, `Cookie`, `Cake`, `Smoothie`, `Juice`
* `Quantity` - number of items bought by a customer.
* `Price Per Unit` - price of each item.
* `Total Spent` - a product of ***`Quantity`*** and ***`Price Per Unit`***.
* `Payment Method` - means of payment. Includes: `Cash`, `Credit Card`, or `Digital Wallet`.
* `Location` - where the customer eats/drinks their order. Could be: `In-Store` or `Takeaway`.
* `Transaction Date` - date of transaction **ONLY** the year `2023`.        

#### Visible Data Quality Issues
The dataset contains common real-world issues, including:
- Missing values/ blanks: the data contains blanks in 7 of the 8 columns.
- Invalid entries, e.g BLANK, ERROR, UNKNOWN
- Incorrect data type

#### License
The data is publicly available on Kaggle for education and portfolio use. To download the latest version of the data, [Click Here](https://www.kaggle.com/datasets/ahmedmohamed2003/cafe-sales-dirty-data-for-cleaning-training).

#### Data Cleaning & Data Preparation

I started by creating a copy of the data in sheet2 (Processing Sheet). I used VBA because it creates an exact copy without changing the data format. The VBA code is as shown below.
~~~ vba
Sub CopyTableToProcessingSheet()
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim tb1 As ListObject
    
    Set ws1 = Sheets("Cafe Sales")
    Set ws2 = Sheets("Processing Sheet")
    
    'Reference table by name
    On Error Resume Next
    Set tb1 = ws1.ListObjects("DirtyCafeSalesData")
    On Error GoTo 0
    
    If tb1 Is Nothing Then
        MsgBox "Table 'DirtyCafeSalesData' not foundon Cafe Sales. ", vbExclamation
        Exit Sub
    End If
    
    ' Copy the data now
    ' Preserve the formats, color, font, borders
    
    tb1.Range.Copy Destination:=ws2.Range("A1")
    
    'Clear clipboard
    Application.CutCopyMode = False
    
    MsgBox "Table copied successfully", vbInformation
End Sub
~~~

##### Data Cleaning

> [!Note]
> This will be a recurring process until I achieve clean data.

A formula that cleans the `Item` column. This formula fills the cells where there is `UNKNOWN`, `ERROR`, or `(blanks)`. 
The formula follows the following steps:
1. Divides `Total Spent` by `Quantity` to fill in the `UNKNOWN`, `ERROR`, or `(blanks)` for the price per unit column. Cells with a value are not calculated.
2. The created variable is then used to fill the item column based on the item price already provided.

* We have two new values in the item column, `Sandwich or Smoothie` and `Cake or Juice`. This is because these products share a price, and I could not determine the exact product. 

~~~ excel
= LET(
             effprice,
                  IFERROR(
                                    IF(OR([@[Price Per Unit]] = "", [@[Price Per Unit]]= "UNKNOWN",[@[Price Per Unit]]= "ERROR"),
                                                 [@[Total Spent]]/[@Quantity],
                                                  [@[Price Per Unit]]),
                                    ""
                                    ),
               IFS(
               NOT(OR([@Item] = "UNKNOWN", [@Item] = "ERROR", [@Item] = "")),
                     [@Item],
               effprice = 2, "Coffee",
              effprice = 1.5, "Tea",
              effprice = 1, "Cookie",
              effprice = 5, "Salad",
              effprice = 4, "Sandwich or Smoothie",
              effprice = 3, "Cake or Juice",
              TRUE,                 ""
        )
     )
~~~ 
