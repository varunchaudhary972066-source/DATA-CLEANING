import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# load data
df = pd.read_excel('sales_data.xlsx')

# clean column names
df.columns = df.columns.str.strip()

# remove TOTAL row if present
df = df[df['Order ID'] != 'TOTAL']

# Part 1 – Data Understanding

print("Shape:", df.shape)
print(df.head(10))
print(df.info())
print(df.isnull().sum())

# Part 2 – Data Cleaning

df.drop_duplicates(inplace=True)

df['Salesperson'] = df['Salesperson'].fillna(df['Salesperson'].mode()[0])
df['Region'] = df['Region'].fillna(df['Region'].mode()[0])
df['Product'] = df['Product'].fillna(df['Product'].mode()[0])
df['Category'] = df['Category'].fillna(df['Category'].mode()[0])
df['Date'] = df['Date'].fillna(df['Date'].mode()[0])

df['Unit Price (₹)'] = df['Unit Price (₹)'].fillna(df['Unit Price (₹)'].mean())
df['Discount (%)'] = df['Discount (%)'].fillna(df['Discount (%)'].mean())

# convert date
df['Date'] = pd.to_datetime(df['Date'], errors='coerce')

# Part 3 – Basic Analysis

print("Average Revenue:", df['Revenue (₹)'].mean())

print("\nHighest Revenue Order:")
print(df.loc[df['Revenue (₹)'].idxmax()])

# Salesperson wise revenue


salesperson_revenue = df.groupby('Salesperson').agg(
    total_revenue=('Revenue (₹)', 'sum'),
    avg_revenue=('Revenue (₹)', 'mean'),
    total_orders=('Order ID', 'count')
)

print("\nSalesperson Revenue:")
print(salesperson_revenue)
salesperson_revenue.to_excel('salesperson_revenue.xlsx')

# Top 5 products by revenue

product_revenue = df.groupby('Product')['Revenue (₹)'].sum().sort_values(ascending=False).head(5)

print("Top 5 Products:")
print(product_revenue)


# Top 5 salespersons


top_salespersons = df.groupby('Salesperson')['Revenue (₹)'].sum().sort_values(ascending=False).head(5)

print("\nTop Salespersons:")
print(top_salespersons)

# Region with highest sales


region_sales = df.groupby('Region')['Revenue (₹)'].sum().sort_values(ascending=False)

print("\nRegion Sales:")
print(region_sales)


# Visualization


# Product vs Revenue
# sns.barplot(x=product_revenue.index, y=product_revenue.values)
# plt.title('Top Products by Revenue')
# plt.xlabel('Product')
# plt.ylabel('Revenue (₹)')
# plt.xticks(rotation=45)
# plt.show()

# Region vs Revenue
# sns.barplot(x=region_sales.index, y=region_sales.values)
# plt.title('Revenue by Region')
# plt.xlabel('Region')
# plt.ylabel('Revenue (₹)')
# plt.show()

# Revenue distribution
plt.hist(df['Revenue (₹)'], bins=10, color='skyblue', edgecolor='black')
plt.title('Revenue Distribution')
plt.xlabel('Revenue (₹)')
plt.ylabel('Frequency')
plt.show()