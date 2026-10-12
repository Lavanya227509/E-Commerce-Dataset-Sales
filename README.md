# E-Commerce-Dataset-Sales - Sales Analysis & Dashboard

## 📌 Overview
This project is an end-to-end analysis of an E-Commerce Sales Dataset from Oct 2023 to Oct 2024. The goal is to clean the data, analyze sales trends, and build an interactive Excel dashboard for business insights.

## 📂 Dataset Structure
- **Customer_Dim:** Customer_ID, Name, Age, Gender, City, State, Country, Loyalty_Level
- **Product_Dim:** Product_ID, Product_Category, Brand, Unit_Price
- **Store_Dim:** Store_ID, State, Region
- **Sales_Fact:** Sales_ID, Customer_ID, Product_ID, Store_ID, Quantity, Discount, Total_Sales_Amount, Order_Date, Payment_Method

Period: Oct 2023 - Oct 2024 | Records: 2000+

## 🎯 Objectives
1. To clean and standardize raw data (e.g., Name formatting using PROPER, TRIM)
2. To integrate multiple sheets using XLOOKUP/VLOOKUP
3. To analyze monthly sales trends and seasonal patterns
4. To identify top-performing product categories and regions
5. To study customer behavior based on loyalty and payment method
6. To create KPIs: Total Sales, Total Orders, AOV
7. To build an interactive Dashboard with Slicers & Timeline
8. To forecast future sales for planning

## 🧹 Data Cleaning Done
- Removed duplicates
- Cleaned `Name` column: `=PROPER(TRIM(CLEAN(B2)))` -> Name2 column
- Handled 'Unknown' values in Loyalty_Level
- Fixed Date formats and Number formats
- Created calculated columns: Total Sales = Qty * Unit Price * (1-Discount)

## 📊 Dashboard Features
- **KPIs:** Total Revenue, Total Orders, Average Sales
- **Charts:** Sales by Month (Line + Forecast), Sales by Category (Bar), Sales by Region (Map), Sales by Loyalty (Pie), Top 5 Products
- **Filters:** Slicers for Region, Category, Loyalty and Timeline for Date

## 🔧 Tools Used
- Microsoft Excel - PivotTable, PivotChart, Slicers, Timeline, Forecast Sheet, XLOOKUP, Conditional Formatting

## 📈 Key Insights
- Electronics and Fashion are the highest revenue categories
- North region leads in total sales
- Platinum loyalty customers contribute higher AOV
- Credit Card is the most used payment method

## 📁 How to Use
1. Download the Excel file
2. Open `Dashboard` sheet
3. Use Slicers to interact

## 👩‍💻 Author
Lavanya227509

---
*Academic Project - Excel for Data Analytics*
