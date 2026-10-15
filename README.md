# Task 15: Product Count Analysis

## 📌 Project Overview

This project analyzes product distribution across categories using the Sample Superstore dataset in Microsoft Excel. It focuses on counting products by category, identifying the category with the highest number of unique products, and presenting the findings in a clear summary table.

## 🎯 Objective

- Count products across different categories.
- Understand the difference between product records and unique products.
- Practice Excel functions such as `COUNTIF` and `COUNTIFS`.
- Identify the category with the highest unique-product count.
- Create a summary table for easy comparison.

## 📂 Dataset

**Dataset:** Sample Superstore

The dataset commonly contains fields such as:
- Order ID
- Product ID
- Product Name
- Category
- Sub-Category
- Sales
- Quantity

The exact columns depend on the dataset version used.

## 🛠️ Tools Used

- Microsoft Excel
- Excel formulas: `COUNTIF`, `COUNTIFS`, `COUNT`, and `COUNTA`
- PivotTables
- Excel charts (optional)

## 📊 Analysis Performed

### 1. Product Count by Category

Counted product records belonging to each category, such as Furniture, Office Supplies, and Technology.

### 2. Unique Product Count

Checked unique products to avoid counting the same product multiple times because it appeared in different sales transactions. Product ID is preferred as the unique identifier when available.

### 3. Largest Category Identification

Compared the product counts across categories and identified the category with the highest unique-product count.

### 4. Summary Table

Created a structured Excel summary containing category names and their corresponding product counts.

## 🧮 Excel Formulas

**Count records by category:**

```excel
=COUNTIF(Orders!C:C,A2)
```

This counts rows matching the category in cell `A2`. Replace the sheet name and column reference with those in your workbook.

**Count unique products in a category (Excel 365):**

```excel
=COUNTA(UNIQUE(FILTER(Orders!F$2:F$10000,(Orders!C$2:C$10000=A2)*(Orders!F$2:F$10000<>""))))
```

This filters products by category, removes repeated product names, and counts the remaining unique names. If available, use Product ID instead of Product Name for more reliable identification.

**Identify the category with the highest count:**

```excel
=INDEX(A2:A4,MATCH(MAX(B2:B4),B2:B4,0))
```

These examples assume the category names are in `A2:A4` and counts are in `B2:B4`. Adjust the ranges to match your workbook.

## 📈 Deliverables

- Product count table by category.
- Identification of the category with the highest unique-product count.
- Excel summary sheet.
- Optional bar chart comparing product counts across categories.

## 💡 Key Learnings

- Using Excel formulas for category-based analysis.
- Understanding the difference between record counts and unique-product counts.
- Summarizing categorical data using PivotTables.
- Comparing product distribution across business categories.
- Presenting findings in a clear and structured format.

## ✅ Conclusion

This task demonstrates how Microsoft Excel can be used to analyze product distribution and summarize category-level information. Counting unique products provides a more accurate view of product variety than simply counting sales rows.

**Internship:** Data Analytics Internship — Veda Technologies
