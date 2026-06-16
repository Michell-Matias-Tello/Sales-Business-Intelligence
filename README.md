# 📊 Sales Data Analytics Dashboard

> **Professional Excel-based Sales Analytics Project** – A comprehensive **Data Analytics** solution for tracking, analyzing, and visualizing sales performance, costs, returns, and trends using **Pivot Tables, Filters, and VBA Macros** in a single Excel workbook.

---

## 📌 **About the Project**

This project is part of the **Data Analytics** initiative in the **Sales** section. It provides a **centralized, interactive Excel dashboard** to monitor key sales metrics, enabling data-driven decision-making.

- **Data Source**: Consolidated sales data (orders, products, customers, salespersons, categories, payment methods, dates, quantities, costs, and returns).
- **Tool**: Microsoft Excel (Pivot Tables, Slicers, Charts, Conditional Formatting, and VBA Macros).
- **Output**: A dynamic **Sales Dashboard** with real-time filtering and visualization capabilities.

---

## ✨ **Features**
   Feature | Description |
 |--------|-------------|
 | **Interactive Dashboard** | Real-time visual representation of sales, costs, returns, and trends. |
 | **Pivot Tables** | Aggregated summaries of sales by person, product, category, payment method, and time. |
 | **Dynamic Filters (Slicers)** | Filter data by year, quarter, category, payment method, and salesperson. |
 | **Trend Analysis** | Visualize sales and cost trends over time (monthly/quarterly). |
 | **Top Performers** | Identify top salespersons, products, and categories. |
 | **Return Tracking** | Monitor returns by category and their impact on revenue. |
 | **Payment Method Breakdown** | Analyze revenue distribution across payment channels. |
 | **VBA Macros** | Automate data updates and filter management. |

---

## 🛠️ **Technologies & Tools**

- **Microsoft Excel** – Core platform for data storage, processing, and visualization.
- **Pivot Tables & Pivot Charts** – Aggregation and dynamic reporting.
- **Slicers** – Interactive filtering.
- **Conditional Formatting** – Highlighting key insights.
- **VBA Macros** – Automation for:
  - **Data Refresh**: Update all pivot tables and charts with new data.
  - **Filter Management**: Clear or apply predefined filter sets.
- **Data Validation** – Ensures consistency in categories, payment methods, and dates.

---
## 🗂️ Project Structure


````

Sales/
├── Sales.xlsx          # Main Excel file containing:
│   ├── Fact_table      # Raw transactional data (orders, sales, returns)
│   ├── Product         # Product master data (ID, name, category, price, cost)
│   ├── Pivot           # PivotTables for aggregated analysis
│   └── Dashboard       # Interactive visual dashboard
└── README.md           # Project documentation (this file)

````
---



## 🔗 Data Relationship Diagram

```mermaid
erDiagram
    Fact_table ||--o{ Product : "contains"
    Fact_table {
        int Order_ID PK
        int Product_ID FK
        int Customer_ID
        int Sales_Person_ID
        int Quantity_Sold
        string Payment_Method
        int Quantity_Returned
        date Order_Date
        decimal Sales_Amount
        decimal Cost_Amount
        decimal Return_Amount
    }
    Product {
        int Product_ID PK
        string Product_Name
        string Category
        decimal Sales_Price
        decimal Cost_Price
    }

````


Key Relationships:

- Each transaction in Fact_table references a product in the Product table via Product_ID.
- Fact_table contains all transactional metrics (sales, costs, returns).
- Product table serves as the dimension table for product details.

---

## 🖼️ **Dashboard Preview**
![Sales Dashboard](Excel-dashboard.png)

### **Key Visualizations**
1. **KPIs Overview**
   - Total Sales: **$2,218,092**
   - Total Cost: **$1,268,584**
   - Total Returns: **$949,508**
   - Units Sold: **249,185**
   - Average Sales per Order: **$110.90**
   - Order Volume: **20,000**

2. **Sales by Salesperson**
   - Bar chart ranking sales performance across 10 salespersons.

3. **Sales and Cost Trend**
   - Line chart tracking monthly sales vs. costs (Jan–Dec).

4. **Average Sales and Returns Trend**
   - Dual-axis chart for average sales and returns over time.

5. **Payment Composition**
   - Pie chart showing distribution by payment method (Cash, Credit Card, Debit Card, Online Payment).

6. **Top 4 Products by Sales**
   - Horizontal bar chart highlighting best-selling products.

7. **Returns by Category**
   - Bar chart for return values across product categories.

---
## 🛠️ **Technologies & Tools**
- **Microsoft Excel** (PivotTables, Slicers, Charts)
- **VBA Macros** (Automation for data refresh and filter management)
- **Power Pivot** (Data modeling for relationships)
- **Conditional Formatting** (Dynamic highlighting)

---
## 🚀 **How to Use**

### **1. Setup**
- Open `Sales.xlsx` in **Microsoft Excel** (2016 or later recommended).
- Enable **Macros** (Required for automation):
  - Go to `File > Options > Trust Center > Trust Center Settings > Macro Settings`.
  - Select **Enable all macros** (for trusted files).

### **2. Data Refresh**
- **Manual Update**:
  - Replace the raw data in the `Fact_table` and `Product` sheets with your dataset.
  - Ensure column headers match the existing structure.
- **Automated Update**:
  - Press `Alt + F8`, select the **`RefreshData`** macro, and click **Run**.
  - *Note: Macros are pre-configured to update PivotTables and charts.*

### **3. Filtering**
- Use the **slicers** (left panel) to filter by:
  - **Payment Method** (Cash, Credit Card, Debit Card, Online Payment)
  - **Year** (2023–2026)
  - **Quarter** (Q1–Q4)
  - **Category** (Soft Drink, Sports Drink, Tea, Water, etc.)
  - **Salesperson** (1–10)

- To **reset all filters**, run the **`ClearFilters`** macro (`Alt + F8`).

### **4. Dashboard Navigation**
- The **Dashboard** sheet auto-updates based on applied filters.
- Hover over charts for detailed tooltips.

---
## 📊 **Data Dictionary**

### **Fact_table**
   Column | Type | Description |
 |--------|------|-------------|
 | `OrderID` | Integer | Unique identifier for each order. |
 | `ProductID` | Integer | Foreign key linking to `Product` table. |
 | `CustomerID` | Integer | Unique identifier for the customer. |
 | `Category` | String | Product category (e.g., Soft Drink, Tea). |
 | `SalesPerson` | String | Identifier for the salesperson. |
 | `QuantitySold` | Integer | Number of units sold. |
 | `PaymentMethod` | String | Payment type (Cash, Credit Card, etc.). |
 | `QuantityReturned` | Integer | Number of units returned. |
 | `OrderDate` | Date | Date of the transaction. |
 | `Month` | String | Month extracted from `OrderDate`. |
 | `Year` | Integer | Year extracted from `OrderDate`. |
 | `Quarter` | String | Quarter (Q1–Q4) extracted from `OrderDate`. |
 | `SalesAmount` | Float | Total sales value (`QuantitySold * SalesPrice`). |
 | `CostAmount` | Float | Total cost value (`QuantitySold * CostPrice`). |
 | `ReturnAmount` | Float | Total return value (`QuantityReturned * SalesPrice`). |

### **Product**
 | Column | Type | Description |
 |--------|------|-------------|
 | `ProductID` | Integer | Unique identifier for the product. |
 | `ProductName` | String | Name of the product. |
 | `Category` | String | Product category. |
 | `SalesPrice` | Float | Selling price per unit. |
 | `CostPrice` | Float | Cost price per unit. |

---
## 🔧 **Macros**
 | Macro Name | Description |
 |------------|-------------|
 | `RefreshData` | Updates all PivotTables, charts, and calculations. |
 | `ClearFilters` | Resets all slicers and filters to default. |

> **Note**: Macros are **password-protected** (if applicable). Contact the project maintainer for access.

---
## 📈 **Key Insights**
- **Top Payment Method**: Cash (31% of transactions).
- **Highest Returns**: Alcoholic Beverage category.
- **Peak Sales Month**: December (highest sales and cost values).
- **Best-Selling Category**: Soft Drinks (highest contribution to total sales).

---
## 🤝 **Contributing**
Contributions are welcome! To contribute:
1. Fork the repository.
2. Create a new branch (`git checkout -b feature/your-feature`).
3. Commit your changes (`git commit -m "Add your feature"`).
4. Push to the branch (`git push origin feature/your-feature`).
5. Open a Pull Request.

---
## 📜 **License**
This project is **proprietary** and intended for internal use. Unauthorized distribution or modification is prohibited.

---
## 📧 **Contact**
For questions or support, refer to the **Sales Analytics** section in **Notion**.
