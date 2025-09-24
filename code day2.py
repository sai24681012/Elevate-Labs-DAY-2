import pandas as pd
import os
import time

# 1. Load the dataset
df = pd.read_excel(
    r"C:\Users\PARAPATLA SAI KUMAR\OneDrive\Desktop\Superstore excel.xlsx",
    engine="openpyxl"
)

# 2. Replace null values with column mean (numeric only)
df = df.fillna(df.mean(numeric_only=True))

# 3. Remove duplicates
df.drop_duplicates(inplace=True)

# 4. Convert date columns
if "Order Date" in df.columns:
    df["Order Date"] = pd.to_datetime(df["Order Date"], errors="coerce")
if "Ship Date" in df.columns:
    df["Ship Date"] = pd.to_datetime(df["Ship Date"], errors="coerce")

# 5. Keep only first 500 rows
df = df.head(500)

# 6. Save cleaned dataset as Excel with timestamp to avoid overwrite issues
timestamp = time.strftime("%H%M%S")
output_file = rf"C:\Users\PARAPATLA SAI KUMAR\OneDrive\Desktop\Superstore_cleaned_{timestamp}.xlsx"

df.to_excel(output_file, index=False, engine="openpyxl")

print(f"✅ Cleaned dataset saved as: {output_file}")
