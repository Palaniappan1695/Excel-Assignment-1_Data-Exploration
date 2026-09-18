# 📊 Excel Product Data Exploration & Analysis

## 📌 Project Overview

This project is a beginner-level **Data Exploration and Analysis project using Microsoft Excel**.

The objective was to analyze a product dataset and apply fundamental Excel functions to summarize numerical data, categorize products based on business rules, perform conditional analysis, and extract structured information from Product IDs.

This project demonstrates my foundational skills in **Excel, data exploration, data transformation, logical analysis, and conditional aggregation**.

---

## 🎯 Project Objectives

The key objectives of this project were to:

* Explore and summarize product data
* Calculate basic statistical measures
* Identify minimum and maximum product prices
* Categorize products based on price
* Perform conditional aggregation
* Count products based on specific conditions
* Extract information from Product IDs
* Build a structured and clearly labeled Excel analysis

---

## 📁 Dataset

The dataset contains information about different products.

### Dataset Attributes

| Column       | Description                        |
| ------------ | ---------------------------------- |
| Product ID   | Unique identifier for each product |
| Product Name | Name of the product                |
| Brand Name   | Brand associated with the product  |
| Quantity     | Quantity of products               |
| Category     | Product category                   |
| Price        | Price of the product               |

The dataset structure and required analysis are defined in the assignment brief.

---

## 🔍 Analysis Performed

### 1. Basic Data Exploration

I used Excel functions to calculate:

* **Total Price** of all products
* **Number of Products**
* **Average Product Price**

Functions used:

```excel
=SUM()
=COUNT()
=AVERAGE()
```

These calculations provide a basic numerical summary of the dataset.

---

### 2. Minimum & Maximum Price Analysis

I identified the lowest and highest product prices using:

```excel
=MIN()
=MAX()
```

This helps understand the overall price range of the products.

The assignment specifically requires identifying both the minimum and maximum price values.

---

### 3. Price Classification Using IF

A new column called **Price Range** was created to classify products based on their price.

Business rule:

| Condition    | Classification |
| ------------ | -------------- |
| Price ≥ $500 | High Price     |
| Price < $500 | Standard Price |

Excel formula:

```excel
=IF(D2>=500,"High Price","Standard Price")
```

This demonstrates the use of **logical conditions for data categorization**.

---

### 4. Conditional Analysis Using SUMIF & COUNTIF

#### Electronics Category

I used `SUMIF` to calculate the total price of products belonging to the **Electronics** category.

```excel
=SUMIF(Category_Range,"Electronics",Price_Range)
```

#### Products Below $100

I used `COUNTIF` to determine the number of products priced below $100.

```excel
=COUNTIF(Price_Range,"<100")
```

These functions demonstrate **conditional aggregation and data filtering**.

---

### 5. Product ID Text Extraction

The Product ID was further analyzed using Excel text functions.

Three new columns were created:

| New Column   | Function  | Purpose                    |
| ------------ | --------- | -------------------------- |
| Day          | `LEFT()`  | Extract first 2 characters |
| Country Code | `RIGHT()` | Extract last 2 characters  |
| Month        | `MID()`   | Extract characters 4–6     |

Example formulas:

```excel
=LEFT(A2,2)
```

```excel
=RIGHT(A2,2)
```

```excel
=MID(A2,4,3)
```

This demonstrates **text manipulation and data transformation** using Excel.

---

## 🛠️ Tools & Skills Used

### Tools

* Microsoft Excel
* GitHub

### Excel Functions

* `SUM`
* `COUNT`
* `AVERAGE`
* `MIN`
* `MAX`
* `IF`
* `SUMIF`
* `COUNTIF`
* `LEFT`
* `RIGHT`
* `MID`

### Data Analysis Skills

* Data Exploration
* Data Summarization
* Logical Analysis
* Conditional Aggregation
* Data Transformation
* Text Manipulation
* Spreadsheet Analysis

The assignment maps these functions to foundational Data Analyst skills such as data summarization, logical analysis, conditional aggregation, text processing, and spreadsheet proficiency.

---

## 📈 Key Learning Outcomes

Through this project, I practiced how to:

* Convert raw product data into useful summaries
* Apply Excel formulas to real-world-style datasets
* Build business-rule-based classifications
* Analyze subsets of data using conditional functions
* Extract meaningful information from structured text
* Organize an analytical worksheet clearly
* Validate formulas and cell references

The assignment emphasizes accurate formulas, correct cell references, clearly labeled columns, and well-organized spreadsheet structure.

---

## 📂 Repository Structure

```text
Excel-Product-Data-Exploration/
│
├── README.md
│
├── data/
│   └── product_dataset.xlsx
│
├── analysis/
│   └── Excel_Assignment_1.xlsx
│
└── screenshots/
    ├── basic_calculations.png
    ├── price_classification.png
    ├── conditional_analysis.png
    └── text_extraction.png
```

> Update the filenames above to match the actual files you upload to your repository.

---

## 📊 Project Workflow

```text
Raw Product Dataset
        ↓
Data Exploration
        ↓
SUM / COUNT / AVERAGE
        ↓
MIN / MAX Analysis
        ↓
Price Classification using IF
        ↓
Conditional Analysis using SUMIF / COUNTIF
        ↓
Product ID Transformation
        ↓
LEFT / RIGHT / MID
        ↓
Final Excel Analysis
```

---

## 🚀 Future Improvements

This project represents my starting point in data analytics. Future versions of this repository may include:

* Excel Pivot Tables
* Interactive Excel Dashboards
* Data Cleaning
* Power Query
* Advanced Excel formulas
* Data Visualization
* SQL-based analysis
* Python-based data analysis
* Business-focused KPI analysis

---

## 👨‍💻 About Me

I am a **budding Data Analyst** developing my skills in Excel, data analysis, SQL, Python, and data visualization.

I am particularly interested in applying data analysis techniques to **finance, business operations, and MIS reporting**.

This repository documents my learning journey and practical projects as I progress toward a career in **Data Analytics**.

---

## 📌 Project Status

**Completed – Beginner Excel Data Analysis Project**

This project was created as part of my learning journey in **Data Analytics – Excel**.
