import pandas as pd
import matplotlib.pyplot as plt

# Step 1: Try reading the Excel file
try:
    df = pd.read_excel("health_data(4).xlsx",sheet_name="Sheet1")
    df.columns = df.columns.astype(str).str.strip()  # clean headers
    print("✅ File loaded. Columns found:", df.columns.tolist())
except Exception as e:
    print("❌ Error loading file:", e)
    exit()

# Step 2: Try to find the real 'Date' column (in case of hidden characters)
for col in df.columns:
    if 'date' in col.lower():  # case-insensitive match
        print(f"🔍 Found date column: '{col}' → Renaming to 'Date'")
        df.rename(columns={col: 'Date'}, inplace=True)

# Step 3: Check again before using
if 'Date' not in df.columns:
    print("❌ 'Date' column NOT found even after cleaning.")
    exit()

# Step 4: Convert to datetime
df['Date'] = pd.to_datetime(df['Date'], errors='coerce')  # avoid crashing


# Calculations
avg_steps = df['Steps'].mean()
total_steps = df['Steps'].sum()
avg_sleep = df['Sleep Hours'].mean()
avg_heart_rate = df['Heart Rate'].mean()
max_hr = df['Heart Rate'].max()
min_hr = df['Heart Rate'].min()
total_calories = df['Calories'].sum()

# Print summary
summary = {
    "Total Days": len(df),
    "Total Steps": int(total_steps),
    "Average Steps": round(avg_steps, 2),
    "Average Sleep (hrs)": round(avg_sleep, 2),
    "Avg Heart Rate": round(avg_heart_rate, 2),
    "Max Heart Rate": max_hr,
    "Min Heart Rate": min_hr,
    "Total Calories Burned": total_calories
}

# Save to new Excel file
summary_df = pd.DataFrame([summary])
summary_df.to_excel("health_report.xlsx", index=False)

print("✅ Report saved as 'health_report.xlsx'.")

# Optional: Chart - Steps Trend
plt.figure(figsize=(10,5))
plt.plot(df['Date'], df['Steps'], marker='o')
plt.title("Daily Steps Trend")
plt.xlabel("Date")
plt.ylabel("Steps")
plt.grid(True)
plt.tight_layout()
plt.savefig("steps_chart.png")
print("📈 Steps chart saved as 'steps_chart.png'.")
