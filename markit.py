I'd be happy to help you combine your fund codes with semicolons using Python! Python is great for this kind of text processing task. Here's a simple script that can help you accomplish this:

```python
import pandas as pd

# Load your Excel file
file_path = "your_file.xlsx"  # Replace with your file path
sheet_name = "Sheet1"  # Replace with your sheet name

# Read the Excel file
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Assuming fund codes are in separate columns
# This will combine all columns with semicolons
df['Combined_Funds'] = df.astype(str).apply(lambda x: ';'.join(x.dropna()), axis=1)

# If fund codes are in a single column but multiple rows
# Uncomment and use this instead:
# combined_funds = ';'.join(df['Fund_Code_Column'].dropna().astype(str).tolist())

# Save the result back to Excel
df.to_excel("result_file.xlsx", index=False)

print("Fund codes combined successfully!")
```

Would you like me to adjust this code to match your specific Excel file structure? For example, do you have your fund codes in multiple columns in a single row, or are they in a single column across multiple rows?