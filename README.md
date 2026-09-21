# Product-Analysis-Excel-Ass1
📊 Data Analytics – Excel Assignment 1
Data Exploration using Microsoft Excel
📌 Assignment Overview

This project is part of the Data Analytics (DA) – Module 1: Excel course.

The objective of this assignment is to perform basic data exploration and analysis using Microsoft Excel. The dataset contains product-related information such as Product ID, Product Name, Brand Name, Quantity, Category, and Price.

📂 Dataset

The Product Dataset contains the following attributes:

Product ID
Product Name
Brand Name
Quantity
Category
Price
🛠️ Excel Functions Used

The following Excel functions were applied in this assignment:

1. Basic Data Exploration
SUM
COUNT
AVERAGE

These functions were used to calculate:

Total price of all products
Number of products
Average product price
2. Minimum and Maximum Analysis
MIN
MAX

These functions were used to identify:

Minimum product price
Maximum product price
3. Logical Function
IF

A new column named Price Range was created.

Classification:

Price ≥ $500 → High Price
Price < $500 → Standard Price
4. Conditional Functions
SUMIF
COUNTIF

These functions were used to:

Calculate the total price of products in the Electronics category
Count products with a price less than $100
5. Text Functions
LEFT
RIGHT
MID

New columns were created from the Product ID:

Day – First 2 characters using LEFT
Country Code – Last 2 characters using RIGHT
Month – 4th to 6th characters using MID
📁 Project Files
Excel File

The Excel workbook contains:

Product dataset
Completed calculations
Applied Excel formulas
Price Range column
Day column
Country Code column
Month column
Clearly labeled results
PDF Documentation

The PDF contains screenshots showing:

Calculated results
Newly created columns
Excel formulas in the formula bar
Completed analysis


🧮 Formula Examples
=SUM(F2:F6)
=COUNT(F2:F6)
=AVERAGE(F2:F6)
=MIN(F2:F6)
=MAX(F2:F6)
=IF(F2>=500,"High Price","Standard Price")
=SUMIF(E2:E6,"Electronics",F2:F6)
=COUNTIF(F2:F6,"<100")
=LEFT(A2,2)
=RIGHT(A2,2)
=MID(A2,4,3)

Note: Cell ranges may vary depending on the dataset used.

🎯 Learning Outcome

Through this assignment, I gained practical experience in using Microsoft Excel for basic data exploration and data transformation.

I learned how to summarize numerical data, apply logical conditions, perform conditional calculations, and extract meaningful information from text fields using Excel functions.
