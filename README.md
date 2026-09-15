
# 📊 Sales Data Analysis & Dashboard

## 📌 Project Overview

This project is an **Excel-based Sales Data Analysis and Dashboard project** created using Microsoft Excel.

The main objective of this project is to analyze sales data, calculate important business metrics, apply Excel formulas and Data Analysis ToolPak features, perform statistical analysis, create Pivot Tables, analyze sales growth, and present the final results through a professional dashboard.

The project demonstrates practical use of Excel for **Data Analysis, Business Intelligence, Statistical Analysis, Data Visualization, and Decision Making**.

---

# 🎯 Project Objectives

The main objectives of this project are:

- Analyze sales and profit performance
- Apply Excel formulas and functions
- Perform customer-wise sales analysis
- Identify top customers based on sales
- Analyze the impact of discounts on profit
- Perform What-If Analysis
- Perform Linear Regression using Data Analysis ToolPak
- Generate Descriptive Statistics
- Analyze monthly sales growth
- Create Pivot Tables for business analysis
- Analyze sales by product category and region
- Create visualizations
- Build a final interactive Sales Dashboard
- Present meaningful business insights and storytelling

---

# 🛠️ Tools & Technologies Used

- Microsoft Excel
- Excel Formulas & Functions
- Conditional Formatting
- What-If Analysis
- Goal Seek / Scenario Analysis
- Data Analysis ToolPak
- Linear Regression
- Descriptive Statistics
- Pivot Tables
- Pivot Charts
- Data Visualization
- Dashboard Design

---

# 📂 Workbook Structure

The workbook contains the following sheets:

1. **Raw Data**
2. **Formulas**
3. **Summary**
4. **Descriptive Statics**
5. **Pivot Table**
6. **Sales Growth**
7. **Visualisation**
8. **Dashboard**

---

# 1️⃣ Raw Data Sheet

The **Raw Data** sheet contains the main sales dataset used for the complete project.

### Columns Included

- Customer_ID
- Customer_Name
- Region
- Product_Category
- Sales
- Quantity
- Discount
- Order_Date
- Profit

The dataset contains **200 analyzed transactions**.

The Raw Data sheet acts as the primary source for formulas, statistical analysis, Pivot Tables, sales growth analysis, and dashboard calculations.

---

# 2️⃣ Formulas Sheet

The **Formulas** sheet contains important Excel formulas used for data analysis.

### Major Excel Functions Used

- INDEX
- MATCH
- LARGE
- SUM
- SUMIF
- NOW
- IF
- COUNTIF
- Other supporting Excel functions

### Top Customer Analysis

The `LARGE` function is used to identify the highest sales values.

The `INDEX` and `MATCH` functions are used to retrieve the corresponding customer names.

Example:

```excel
=INDEX('Raw Data'!B:B,MATCH(C2,'Raw Data'!E:E,0))
````

This helps identify customers associated with high sales values.

---

# 3️⃣ Summary Sheet

The **Summary** sheet contains the output of the Linear Regression analysis.

The regression model analyzes the relationship between **Quantity and Profit**.

### Regression Results

* Multiple R: approximately **0.7735**
* R Square: approximately **0.5983**
* Adjusted R Square: approximately **0.5962**
* Standard Error: approximately **86.33**
* Observations: **200**

The R Square value of approximately 0.598 indicates that around **59.8% of the variation in the dependent variable is explained by the regression model**.

The regression coefficient for the X Variable is approximately **0.3511**, indicating a positive relationship between the variables used in the model.

---

# 4️⃣ Descriptive Statics Sheet

The **Descriptive Statics** sheet contains statistical analysis generated using the **Data Analysis ToolPak**.

The analysis is performed for:

* Sales
* Quantity
* Discount
* Profit

### Statistical Measures Included

* Mean
* Standard Error
* Median
* Mode
* Standard Deviation
* Sample Variance
* Kurtosis
* Skewness
* Range
* Minimum
* Maximum
* Sum
* Count

### Important Results

**Sales**

* Mean: approximately ₹976.09
* Median: approximately ₹976.72
* Standard Deviation: approximately ₹299.27
* Minimum: ₹258.51
* Maximum: ₹1,923.66
* Total Sales: approximately ₹195,217.76
* Count: 200

**Quantity**

* Mean: approximately 10
* Median: 10
* Minimum: 1
* Maximum: 19
* Total Quantity: 1,999

**Discount**

* Mean: approximately 9.73%
* Median: 10%
* Minimum: 0%
* Maximum: 20%

**Profit**

* Mean: approximately ₹341.85
* Median: approximately ₹332.69
* Standard Deviation: approximately ₹135.86
* Minimum: ₹64.63
* Maximum: ₹769.46
* Total Profit: approximately ₹68,369.90

---

# 5️⃣ Pivot Table Sheet

The **Pivot Table** sheet contains multiple Pivot Table analyses.

Pivot Tables are used to summarize and analyze the dataset from different perspectives.

### Product Category Analysis

Sales are analyzed across:

* Books
* Clothing
* Electronics
* Furniture
* Office Supplies

### Regional Analysis

Sales are analyzed across:

* Central
* East
* North
* South
* West

### Monthly Sales Analysis

The Pivot Table also provides sales information by month.

This makes it easier to identify changes in sales performance throughout the year.

### Product & Region Analysis

The Pivot Table also compares product categories across different regions.

This helps identify which products perform better in specific regions.

---

# 6️⃣ Sales Growth Sheet

The **Sales Growth** sheet is used to analyze sales performance over time.

The analysis uses the `Order_Date` field and summarizes sales values according to the available date/month information.

This analysis helps identify:

* Sales trends
* Monthly performance
* Growth patterns
* High-performing periods
* Low-performing periods

Conditional Formatting and visual indicators can be used to make increases and decreases easier to identify.

---

# 7️⃣ Visualisation Sheet

The **Visualisation** sheet is designed for presenting analytical results visually.

The project uses charts to make the data easier to understand.

### Main Visualizations

#### 📊 Bar Chart

Used to compare sales across product categories.

It helps identify which product categories generate higher sales.

#### 📈 Line Chart

Used to display sales trends over time.

It helps identify increases and decreases in sales performance across months.

#### 🥧 Pie Chart

Used to represent the distribution of sales across regions or categories.

It provides a quick visual comparison of each region's contribution.

---

# 8️⃣ Dashboard Sheet

The **Dashboard** is the final presentation layer of the project.

It brings together important sales information and visualizations into one place.

The dashboard provides a quick overview of business performance.

### Dashboard Includes

* Sales KPIs
* Product performance
* Regional performance
* Sales trends
* Charts
* Visual indicators
* Analytical summaries

The dashboard is designed to make the analysis easy to understand for users and decision-makers.

---

# 📈 Key Business Analysis

The project provides several important business insights.

### Product Performance

Books generated the highest total sales among the product categories, followed by Clothing and Electronics.

### Regional Performance

Central and West are among the strongest regions in terms of total sales, while East has comparatively lower sales.

### Profit Performance

The dataset generated approximately **₹68,369.90 total profit**.

### Sales Performance

The dataset generated approximately **₹195,217.76 total sales** from 200 transactions.

### Quantity Performance

A total of **1,999 units** were sold.

---

# 🔎 Statistical Analysis

Descriptive Statistics were used to understand the distribution and characteristics of the dataset.

The analysis provides information about:

* Central tendency
* Data variation
* Distribution
* Minimum and maximum values
* Overall totals

This helps convert raw data into meaningful statistical information.

---

# 📉 Linear Regression Analysis

Linear Regression was performed using the **Excel Data Analysis ToolPak**.

The purpose of regression analysis is to understand the relationship between the selected variables.

### Regression Summary

* Observations: 200
* Multiple R: 0.7735
* R Square: 0.5983
* Adjusted R Square: 0.5962
* Standard Error: 86.33

The model indicates a **positive relationship** between the variables used in the regression analysis.

---

# 🎯 What-If Analysis

What-If Analysis is used to understand how changing one input can affect the final business result.

The project includes discount scenarios to analyze the effect of different discount levels.

### Discount Scenarios

The analysis considers discount values such as:

* 5%
* 10%
* 15%
* 20%

This helps understand how changing discount levels can influence sales/profit outcomes.

What-If Analysis is useful for business planning and decision-making.

---

# 👥 Customer Analysis

Customer-level analysis is performed using Excel formulas.

The project uses functions such as:

* LARGE
* INDEX
* MATCH
* SUMIF

These functions help identify customers with high sales values and summarize customer-level performance.

---

# 📊 Important Excel Concepts Used

The project demonstrates practical knowledge of:

* Data Cleaning
* Excel Formulas
* INDEX & MATCH
* LARGE
* SUMIF
* COUNTIF
* NOW
* Conditional Formatting
* What-If Analysis
* Scenario Analysis
* Goal Seek
* Data Analysis ToolPak
* Linear Regression
* Descriptive Statistics
* Pivot Tables
* Pivot Charts
* Data Visualization
* Dashboard Development
* Business Storytelling

---

# 💡 Key Findings

Based on the analysis:

1. Total sales are approximately **₹195,217.76**.
2. Total profit is approximately **₹68,369.90**.
3. Total quantity sold is **1,999 units**.
4. The dataset contains **200 transactions**.
5. Books is the highest-performing product category by sales.
6. Central and West are among the strongest regions by sales.
7. The average sales value is approximately **₹976.09** per transaction.
8. The average quantity per transaction is approximately **10 units**.
9. The average discount is approximately **9.73%**.
10. The regression analysis shows a positive relationship between the selected variables.
11. Descriptive Statistics provide detailed information about sales, quantity, discount, and profit distribution.
12. Pivot Tables provide a flexible way to compare product, region, month, sales, quantity, and profit performance.

---

# 📚 Learning Outcomes

Through this project, I learned how to:

* Work with real-world style sales data
* Organize and analyze large datasets
* Use Excel formulas effectively
* Perform statistical analysis
* Use Data Analysis ToolPak
* Perform Linear Regression
* Generate Descriptive Statistics
* Create and analyze Pivot Tables
* Analyze sales growth
* Create meaningful charts
* Build an Excel Dashboard
* Present business insights using data storytelling
* Use Excel for business decision-making

---

# 📤 Project Output

The final project provides a complete Excel-based Sales Analytics solution consisting of:

* Raw Dataset
* Formula-Based Analysis
* Regression Analysis
* Descriptive Statistics
* Pivot Table Analysis
* Sales Growth Analysis
* Data Visualizations
* Final Dashboard

The project demonstrates how raw sales data can be transformed into **meaningful analytical information and business insights using Microsoft Excel**.

---

# 🏁 Conclusion

This project demonstrates the practical use of Microsoft Excel as a **Data Analysis and Business Intelligence tool**.

Starting from raw transaction data, the project applies formulas, statistical techniques, What-If Analysis, regression, Pivot Tables, sales growth analysis, visualizations, and dashboard design.

The final Dashboard provides a clear and concise overview of sales performance and helps users understand important business trends and insights.

Overall, this project demonstrates practical skills in **Excel Data Analysis, Statistical Analysis, Data Visualization, Dashboard Creation, and Business Storytelling**.

---

# 👨‍💻 Author

**Sarth Thakar**
