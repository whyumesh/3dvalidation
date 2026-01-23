import pandas as pd
import matplotlib.pyplot as plt

# =========================
# CONFIG
# =========================
FILE_PATH = "Book29.xlsx"   # change if needed
TOP_N = 10

# =========================
# LOAD DATA
# =========================
df = pd.read_excel(FILE_PATH)

print("Data Loaded Successfully")
print(f"Rows: {df.shape[0]}, Columns: {df.shape[1]}")

# =========================
# BASIC CLEANING
# =========================
df.columns = df.columns.str.strip()

# Normalize key columns
df['Status'] = df['Status'].fillna('Unknown')
df['TAT'] = df['TAT'].fillna('Pending')
df['State'] = df['State'].fillna('Unknown')
df['Division'] = df['Division'].fillna('Unknown')

# =========================
# BUSINESS KPIs
# =========================
total_cases = len(df)
done_cases = (df['Status'] == 'Done').sum()
outside_tat = (df['TAT'] == 'Outside TAT').sum()
pending_cases = (df['TAT'] == 'Pending').sum()

print("\n===== BUSINESS KPIs =====")
print(f"Total Cases        : {total_cases}")
print(f"Completed (Done)   : {done_cases}")
print(f"Outside TAT        : {outside_tat}")
print(f"Pending            : {pending_cases}")

# =========================
# VISUALIZATION 1: STATUS
# =========================
plt.figure()
df['Status'].value_counts().plot(kind='bar')
plt.title("Employee Status Distribution")
plt.xlabel("Status")
plt.ylabel("Count")
plt.show()

# =========================
# VISUALIZATION 2: TAT
# =========================
plt.figure()
df['TAT'].value_counts().plot(kind='bar')
plt.title("Turnaround Time (TAT) Distribution")
plt.xlabel("TAT Category")
plt.ylabel("Count")
plt.show()

# =========================
# VISUALIZATION 3: TOP STATES
# =========================
plt.figure()
df['State'].value_counts().head(TOP_N).plot(kind='bar')
plt.title(f"Top {TOP_N} States by Employee Count")
plt.xlabel("State")
plt.ylabel("Employees")
plt.show()

# =========================
# VISUALIZATION 4: DIVISION LOAD
# =========================
plt.figure()
df['Division'].value_counts().plot(kind='bar')
plt.title("Employees by Division")
plt.xlabel("Division")
plt.ylabel("Count")
plt.xticks(rotation=45, ha='right')
plt.show()

# =========================
# VISUALIZATION 5: PROCESSING DELAY
# =========================
plt.figure()
df['Diff'].plot(kind='hist', bins=20)
plt.title("Processing Time Difference (Diff)")
plt.xlabel("Days")
plt.ylabel("Frequency")
plt.show()

# =========================
# EXCEPTION REPORT
# =========================
exceptions = df[
    (df['TAT'] == 'Outside TAT') | (df['TAT'] == 'Pending')
]

print("\n===== EXCEPTION SUMMARY =====")
print(exceptions[['Emp Name', 'State', 'Division', 'TAT', 'Remark']].head(10))

print("\nScript Execution Completed")
