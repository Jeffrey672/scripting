import pandas as pd

# Load Excel file
excel_file_path = 'C:\Users\JeffreyOs\Downloads\277 Front Desktops.xlsx'
df_excel = pd.read_excel(excel_file_path)

# Preview the data
print("Preview of Excel data:")
print(df_excel.head())

# Extract serial numbers from the Excel file
excel_serials = df_excel['Serial Number'].astype(str).str.strip().tolist()

# Load SCCM data (replace with your actual SCCM file path)
# df_sccm = pd.read_csv('sccm_serials.csv')  # Uncomment when you have SCCM data
# sccm_serials = df_sccm['Serial Number'].astype(str).str.strip().tolist()

# Simulated SCCM serials (replace with real data)
sccm_serials = ['ABC123', 'XYZ789', 'DEF456', 'GHI000']  # Example SCCM data

# Convert lists to sets for comparison
set_excel = set(excel_serials)
set_sccm = set(sccm_serials)

# Find serials in both Excel and SCCM
both_serials = set_excel.intersection(set_sccm)

# Output the comparison
# print("\nSerial numbers only in Excel:")
# print(both_serials)

# print("\nSerial numbers only in SCCM:")
# print(only_in_sccm)

print(f"Total serials in both Excel and SCCM: {len(both_serials)}")