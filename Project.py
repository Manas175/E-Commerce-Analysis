import pandas as pd # type: ignore
import numpy as np # type: ignore
import matplotlib.pyplot as plt # type: ignore

#1: Create sample data
np.random.seed(1)
data = {
    'Order ID': range(1, 51),
    'Product': np.random.choice(['Phone', 'Shirt', 'Blender', 'Lipstick'], 50),
    'Category': np.random.choice(['Electronics', 'Clothing', 'Home', 'Beauty'], 50),
    'Price': np.random.uniform(10, 200, 50).round(2),
    'Quantity': np.random.randint(1, 10, 50),
    'Order Date': pd.date_range(start='2024-01-01', periods=50, freq='D')
}
df = pd.DataFrame(data)
df['Revenue'] = df['Price'] * df['Quantity']

#2: Finding missing values and duplicates
df.loc[5, 'Price'] = np.nan
df.loc[10, 'Quantity'] = np.nan
df = pd.concat([df, df.iloc[[2]]], ignore_index=True)  # Add duplicate

#3: Checking for missing values and duplicates
print("Missing values:\n", df.isnull().sum())
print("Duplicates:", df.duplicated().sum())

#4: Cleaning the data
df_cleaned = df.dropna().drop_duplicates()

#5: Creating pivot table (Revenue)
pivot = df_cleaned.pivot_table(index='Category', values='Revenue', aggfunc='sum')
print("\nRevenue by Category:\n", pivot)

#6: Creating all 3 charts
# Bar Chart - Revenue by Category
pivot.plot(kind='bar', title='Revenue by Category')
plt.tight_layout()
plt.savefig('bar_chart.png')
plt.close()

# Line Chart - Revenue over Time
revenue_by_date = df_cleaned.groupby('Order Date')['Revenue'].sum()
revenue_by_date.plot(kind='line', title='Revenue Over Time')
plt.tight_layout()
plt.savefig('line_chart.png')
plt.close()

# Pie Chart - Revenue by Category
pivot['Revenue'].plot(kind='pie', autopct='%1.1f%%', title='Revenue Distribution')
plt.ylabel('')
plt.tight_layout()
plt.savefig('pie_chart.png')
plt.close()

#7: Save to Excel
with pd.ExcelWriter('ecommerce_project.xlsx') as writer:
    df.to_excel(writer, sheet_name='Raw Data', index=False)
    df_cleaned.to_excel(writer, sheet_name='Cleaned Data', index=False)
    pivot.to_excel(writer, sheet_name='Pivot Table')
