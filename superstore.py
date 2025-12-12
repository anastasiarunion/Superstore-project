import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import sqlite3

# Load Excel dataset
df = pd.read_excel("sample_-_superstore.xls")
df['Order Date'] = pd.to_datetime(df['Order Date'])

# Filter data for 2014
df_2014 = df[df['Order Date'].dt.year == 2014]

# Aggregate data
sales_by_category = df_2014.groupby('Category')['Sales'].sum().reset_index()
profit_by_category = df_2014.groupby('Category')['Profit'].sum().reset_index()
monthly_sales = df_2014.groupby(df_2014['Order Date'].dt.month)[
    'Sales'].sum().reset_index()
profit_by_region = df.groupby('Region')['Profit'].sum()

# Export aggregated data for Tableau
sales_by_category.to_csv("sales_by_category.csv", index=False)
df.to_csv("superstore_cleaned.csv", index=False)

# Visualizations
plt.figure(figsize=(8, 5))
sns.barplot(x=sales_by_category.index, y=sales_by_category.values)
plt.title('Total Sales by Category')
plt.ylabel("Sales ($)")
plt.show()

profit_by_region.plot(kind='pie', autopct='%1.1f%%',
                      figsize=(6, 6), title="Profit by Region")
plt.ylabel("")
plt.show()

sales_per_year = df.groupby(df['Order Date'].dt.year)['Sales'].sum()
sales_per_year.plot(kind='line', marker='o', title="Sales Over Years")
plt.ylabel("Sales ($)")
plt.show()

# SQL demonstration using SQLite
conn = sqlite3.connect('superstore.db')
df = pd.read_csv("superstore_cleaned.csv")
df.to_sql('superstore', conn, index=False, if_exists='replace')

# SQL queries
query_category = """
SELECT Category, SUM(Sales) AS total_sales, SUM(Profit) AS total_profit
FROM superstore
WHERE substr("Order Date", 1, 4) = '2014'
GROUP BY Category
"""
result_category = pd.read_sql(query_category, conn)
print(result_category)

query_monthly = """
SELECT strftime('%m', "Order Date") AS month, SUM(Sales) AS monthly_sales 
FROM superstore 
WHERE substr("Order Date", 1, 4) = '2014' 
GROUP BY month 
ORDER BY month
"""
monthly_sales = pd.read_sql(query_monthly, conn)
print(monthly_sales)

query_region = """
SELECT Region, SUM(Profit) AS total_profit
FROM superstore
GROUP BY Region
"""
profit_region = pd.read_sql(query_region, conn)
print(profit_region)

# Export SQL results
monthly_sales.to_csv("monthly_sales.csv", index=False)
profit_region.to_csv("profit_by_region.csv", index=False)
