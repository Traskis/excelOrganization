import pandas as pd

# 1. Load the Excel file
file_name = "Employees.xlsx"
df = pd.read_excel(file_name)

# 2. Count employees per department with clear columns
employees_per_department = df["Department"].value_counts().reset_index()
employees_per_department.columns = ["Department", "Count"]

# 3. Find the highest paid employee
highest_paid = df.loc[df["Salary"].idxmax()]

# 4. Filter employees with salary over 3000
high_earners = df[df["Salary"] > 3000]

# 5. Save results to a new Excel file
with pd.ExcelWriter("Analysis_Results.xlsx", engine="openpyxl") as writer:
    employees_per_department.to_excel(writer, sheet_name="Department Stats", index=False)
    pd.DataFrame([highest_paid]).to_excel(writer, sheet_name="Top Earner", index=False)
    high_earners.to_excel(writer, sheet_name="Salary > 3000", index=False)

print("✅ Analysis complete! Results saved in 'Analysis_Results.xlsx'")
