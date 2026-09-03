# Data Cleaning and Visualization of Cafe Sales Data

### Project Overview

Real-world datasets are rarely clean, and messy data is one of the common challenges data analysts face before drawing any meaningful insights. In this project, I work with a "Dirty Cafe Sales" dataset, a sales record containing common real-world data quality issues like missing values, inconsistency in formatting, and incorrect data types.

The main objective of this project is to clean and prepare the dataset while minimizing data loss(data deletion). Normally, blank columns are often dropped, a mistake that can be avoided in data cleaning, and one of the goals here is to demonstrate techniques that preserve as much data as possible.

Two common imputation methods, mean imputation and median imputation, are going to be the key focus of this project. By imputing blank values using these methods, I am going to evaluate how each method affects the overall distribution, summary statistics, and visual outcomes of the sales data. This comparison aims to highlight the practical differences between the two approaches and build a stronger understanding of which method is more appropriate. 

The cleaned data will then be explored through EDA to uncover patterns and trends in the cafe sales, such as the best-selling item, sales trends in the year 2023, and revenue distribution. These reports will then be brought together into an interactive Excel dashboard, allowing stakeholders to visually explore cafe sales performance at a glance.

This project in it entirity, from data cleaning to the interactive dashboard, will be executed exclusively in **Microsoft Excel**, showcasing not only how far Excel's built-in tools can be pushed to handle a full structure data analysis workflow, but also my expertise in Excel formulas, PivotTables, PivotCharts, conditional formatting, and data validation.

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

#### Visible Data Quality Issues
The dataset contains common real-world issues, including:
- Missing values/ blanks- the data contains blanks in 7 of the 8 columns.
- Invalid entries, e.g BLANK, ERROR, UNKNOWN
- Incorrect data type

#### License
The data is publicly available on Kaggle for education and portfolio use. To download the latest version of the data and read the column descriptions, [Click Here](https://www.kaggle.com/datasets/ahmedmohamed2003/cafe-sales-dirty-data-for-cleaning-training).   
