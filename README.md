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

## 📁 **Project Structure**


Sales.xlsx
├── Fact_table          # Raw transactional data (Order ID, Product, Customer, Sales Person, Quantity, Payment Method, etc.)
├── Product             # Product master data (Product ID, Name, Category, Sales Price, Cost Price)
├── Pivot               # PivotTables for aggregated analysis (Sales, Costs, Returns by Salesperson, Category, etc.)
└── Dashboard           # Interactive dashboard with visualizations and KPIs
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

Each transaction in Fact_table references a product in the Product table via Product_ID.
Fact_table contains all transactional metrics (sales, costs, returns).
Product table serves as the dimension table for product details.


📊 Dashboard Preview
The dashboard includes the following visualizations and metrics:


  
    
      Metric
      Description
    
  
  
    
      Total Sales
      $2,218,091.58 (2023–2026)
    
    
      Total Cost
      $1,268,584.00
    
    
      Total Returns
      $949,508.00
    
    
      Units Sold
      249,185
    
    
      Average Sales
      $110.90 per order
    
    
      Order Volume
      20,000 orders
    
  



Visualizations

Sales by Salesperson: Bar chart ranking sales performance.
Sales and Cost Trend: Line chart showing monthly trends.
Average Sales & Returns Trend: Dual-axis line chart for comparative analysis.
Payment Composition: Pie chart of payment method distribution.
Top 4 Products by Sales: Horizontal bar chart of best-selling products.
Returns by Category: Bar chart of return values per product category.

🛠️ Technologies & Tools

Microsoft Excel (PivotTables, Charts, Slicers, Conditional Formatting)
VBA (Visual Basic for Applications) (Macros for automation)
Data Validation (Dropdown filters for dynamic selection)

🚀 Setup & Usage
Prerequisites

Microsoft Excel (2016 or later recommended for full macro support).
Enable Macros in Excel settings (required for automation).
How to Use


Open the File:

Launch Sales.xlsx in Excel.
Enable macros when prompted.


Refresh Data:

Press Alt + F8, select the "Refresh_Data" macro, and click Run.
This updates all PivotTables and charts with the latest data.


Reset Filters:

Press Alt + F8, select the "Clear_Filters" macro, and click Run.
This removes all applied filters to restore the default view.


Interact with the Dashboard:

Use the slicers on the left to filter by:

Payment Method (Cash, Credit Card, Debit Card, Online Payment)
Year (2023, 2024, 2025, 2026)
Quarter (Q1, Q2, Q3, Q4)
Category (Soft Drink, Sports Drink, Tea, Water, Alcoholic Beverage, Coffee, Juice)



🤖 Macros


  
    
      Macro Name
      Purpose
      Shortcut
    
  
  
    
      Refresh_Data
      Updates all PivotTables, charts, and calculations with the latest data.
      Alt + F8 → Run
    
    
      Clear_Filters
      Removes all applied filters to reset the dashboard view.
      Alt + F8 → Run
    
  




Note: Macros are password-protected. Ensure macros are enabled in Excel settings.


📈 Key Insights

Cash is the most used payment method, accounting for the largest share of total orders.
Total returns of $949.5K present a significant opportunity for improvement in product quality or customer satisfaction.
Top products contribute evenly to total sales, suggesting a balanced product portfolio.
Sales trends show seasonal fluctuations, with peaks in specific quarters.

🤝 Contributing
Contributions are welcome! To contribute:

Fork the repository.
Create a new branch (git checkout -b feature/your-feature).
Commit your changes (git commit -m "Add: Your feature description").
Push to the branch (git push origin feature/your-feature).
Open a Pull Request.

Note: Ensure all macros are documented and tested before submission.


📜 License
This project is proprietary and intended for internal use within the Data Analytics team. Unauthorized distribution or modification is prohibited.

📧 Contact
For questions or support, refer to the Sales Analytics section in Notion.

