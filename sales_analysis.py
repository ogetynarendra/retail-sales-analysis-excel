import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# 📥 Load the dataset (fixing encoding issue)
df = pd.read_csv('SampleSuperstore.csv', encoding='ISO-8859-1')

# 📌 Basic Info
print("First 5 Rows of the Data:")
print(df.head())
print("\nData Info:")
print(df.info())
print("\nMissing Values:")
print(df.isnull().sum())

# 🧹 Clean the Data
df.drop_duplicates(inplace=True)

# Optional: Drop columns that aren't useful
df.drop(['Row ID', 'Postal Code'], axis=1, inplace=True)

# 🎯 1. Total Sales by Category
category_sales = df.groupby('Category')['Sales'].sum().sort_values()
category_sales.plot(kind='barh', title='Total Sales by Category', color='skyblue')
plt.xlabel("Total Sales")
plt.ylabel("Category")
plt.tight_layout()
plt.show()

# 🎯 2. Sales vs Profit by Region
region_profit = df.groupby('Region')[['Sales', 'Profit']].sum().sort_values(by='Sales')
region_profit.plot(kind='bar', title='Sales vs Profit by Region', figsize=(8, 5))
plt.ylabel("Amount")
plt.tight_layout()
plt.show()

# 🎯 3. Top 10 Products by Sales
top_products = df.groupby('Product Name')['Sales'].sum().sort_values(ascending=False).head(10)
top_products.plot(kind='bar', title='Top 10 Products by Sales', figsize=(10, 5), color='green')
plt.ylabel("Sales")
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.show()

# 🎯 4. Profit by Sub-Category
subcategory_profit = df.groupby('Sub-Category')['Profit'].sum().sort_values()
subcategory_profit.plot(kind='barh', title='Profit by Sub-Category', figsize=(10, 6), color='orange')
plt.xlabel("Total Profit")
plt.ylabel("Sub-Category")
plt.tight_layout()
plt.show()

# 🎯 5. Discount vs Profit Scatter Plot
plt.figure(figsize=(8, 5))
sns.scatterplot(x='Discount', y='Profit', data=df, color='red')
plt.title("Discount vs Profit")
plt.xlabel("Discount")
plt.ylabel("Profit")
plt.tight_layout()
plt.show()
# 📤 Export grouped data to Excel
with pd.ExcelWriter("sales_analysis_output.xlsx", engine="openpyxl") as writer:
    # Exporting total sales by category
    category_sales.to_frame().to_excel(writer, sheet_name="Sales_by_Category")

    # Exporting sales and profit by region
    region_profit.to_excel(writer, sheet_name="Sales_Profit_by_Region")

    # Exporting top products by sales
    top_products.to_frame().to_excel(writer, sheet_name="Top_Products_by_Sales")

    # Exporting profit by sub-category
    subcategory_profit.to_frame().to_excel(writer, sheet_name="Profit_by_SubCategory")
