<img width="971" height="569" alt="image" src="https://github.com/user-attachments/assets/cf0b3ff3-c9ea-4d32-8731-1e68fd2aedd3" /># Excel Data Analysis Project


## Objective

This project presents an interactive Coffee Sales Dashboard built in Excel to analyze sales performance and customer behavior.
The dashboard should present a dynamic visualisation, including sales trend, sales by country and product type, enabling users to filter data by date ranges and other critical dimensions.
The project must transform raw data into the one that helps business fascilitate data-driven decisions.


### Skills Learned

Data Wrangling / Preprocessing
- Cleaned and transformed raw data into a structured, tidy format using Excel functions like XLOOKUP and INDEX

Data Formatting / Analysis
- Ensured data accuracy and clarity by implementing consistent formats to resolve potential errors
- Created calculated columns by using excel formulas to uncover key business insights

Interactive Data Visualization
- Built dynamic dashboards using PivotTables and PivotCharts to effectively communicate trends and patterns
- Integrating Slicers and Timelines, enabling intuitive filtering and exploration of the data


### Tools Used

Microsoft Excel


## Steps

First, I downloaded the fictitious coffee sales data from a [public github account]([https://www.google.com](https://github.com/mochen862/excel-project-coffee-sales/tree/main)) that provides datasets for excel practices

### 1. Data integration: Combining sheets

Taking a first look at the coffee sales data, I realised that there are three sheets, which are orders, customers and products, in excel.

- Used XLOOKUP to pull in customer details, name, email and country
- Used INDEX-MATCH formula to efficiently retrieve all product information (coffee type, roast, size, price).

### 2. Data Enrichment: Creating "Sales" column

Since I wanted to show the sales trend and compare sales for different countries/customers, I needed to create a "sales" column

- Created a Sales column by multiplying Unit Price by Quantity

### 3. Data Transformation and Formatting: Changing data into clear and the most relevent form

I ensured that all data are in the most understandable formats, and are not prone to any potential errors

- Ensured data integrity by checking for and removing duplicate records

- Used nested IF statements to create new columns with full names for
  - coffee types (Arabica, Robusta, etc.)
  - roast levels (Light, Medium, Dark) from their abbreviated codes

- Applied custom formatting for clarity
  - dates were formatted as DD-MMM-YYYY e.g. 05/Sep/2019 (to prevent confusions by UK/US formats)
  - package sizes were labeled with "kilo", and currency values were formatted as US Dollars

### 4. Data anlysis: Using PivotTable and Pivotchart to summarize key business metrics

I created a PivotTable for the following metrics, so that these can be used to create charts that give meaningful insights 

- Total Sales Over Time: A line chart showing monthly sales trends, broken down by coffee type

<img width="465" height="370" alt="image" src="https://github.com/user-attachments/assets/a09806e4-0887-4720-804b-f1c9c26b1007" />

- Sales by Country: A bar chart comparing revenue across different countries

<img width="502" height="168" alt="image" src="https://github.com/user-attachments/assets/504290e5-ffdb-4e1d-b1e2-2c23f8a13ff7" />

- Top 5 Customers: A bar chart identifying the customers with the highest value (Based on how much they have paid in total. More than 5 customers can get displayed if they had the same value)

<img width="504" height="195" alt="image" src="https://github.com/user-attachments/assets/f9f44c8e-94a9-46d9-9b23-f7f7111d7258" />

### 5. Interactive Dashboard Development

I designed a cohesive Coffee Sales Dashboard on a dedicated worksheet.

- Integrated Slicers for Loyalty Card, Roast Type, and Size and a Timeline (for Order Date) to enable user-friendly data filtering

<img width="973" height="600" alt="image" src="https://github.com/user-attachments/assets/a32a44eb-1124-49b7-943c-fa693152d016" />
