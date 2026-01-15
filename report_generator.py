import pandas as pd
import matplotlib.pyplot as plt

# -----------------------------
# 1. Load Data
# -----------------------------
def load_data(file_path="data.csv"):
    df = pd.read_csv(file_path, encoding="latin1")
    df.rename(columns=lambda x: x.strip(), inplace=True)  # remove spaces
    required_cols = ["Sales", "Profit", "Category", "Segment"]
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"Missing required column: {col}")
    return df

# -----------------------------
# 2. Clean Data
# -----------------------------
def clean_data(df):
    df.drop_duplicates(inplace=True)
    df[["Sales","Profit"]] = df[["Sales","Profit"]].fillna(0)
    df["Category"] = df["Category"].fillna("Unknown")
    df["Segment"] = df["Segment"].fillna("Unknown")
    df = df[df["Sales"] >= 0]
    return df

# -----------------------------
# 3. Calculate KPIs
# -----------------------------
def calculate_kpis(df):
    total_sales = df["Sales"].sum()
    total_profit = df["Profit"].sum()
    profit_margin = (total_profit / total_sales * 100) if total_sales != 0 else 0
    avg_order_value = df["Sales"].mean()

    kpi_df = pd.DataFrame({
        "Metric": ["Total Sales", "Total Profit", "Profit Margin (%)", "Average Order Value"],
        "Value": [round(total_sales,2), round(total_profit,2), round(profit_margin,2), round(avg_order_value,2)]
    })

    # Format numbers for Excel readability
    kpi_df["Value"] = kpi_df["Value"].apply(lambda x: f"{x:,.2f}")
    return kpi_df

# -----------------------------
# 4. Category Analysis
# -----------------------------
def analyze_categories(df):
    category_summary = df.groupby("Category").agg(
        Total_Sales=("Sales","sum"),
        Total_Profit=("Profit","sum")
    ).reset_index()

    category_summary["Profit_Margin (%)"] = (category_summary["Total_Profit"] / category_summary["Total_Sales"] * 100)

    # Flag high revenue but low profit categories (<20% margin)
    category_summary["High_Rev_Low_Profit"] = category_summary["Profit_Margin (%)"] < 20

    flagged = category_summary[category_summary["High_Rev_Low_Profit"]]
    if not flagged.empty:
        print("⚠ Categories with high revenue but low profit margin:\n", 
              flagged[["Category","Total_Sales","Total_Profit","Profit_Margin (%)"]])
    return category_summary

# -----------------------------
# 5. Generate Charts
# -----------------------------
def plot_category_chart(category_summary):
    ax = category_summary.plot(
        x="Category",
        y=["Total_Sales","Total_Profit"],
        kind="bar",
        figsize=(8,5),
        color=["#1f77b4","#ff7f0e"],
        title="Sales vs Profit by Category"
    )
    ax.set_ylabel("Amount ($)")
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig("sales_profit_by_category.png")
    plt.close()

# -----------------------------
# 6. Export Excel
# -----------------------------
def export_excel(df, kpi_df, category_summary):
    with pd.ExcelWriter("output_report.xlsx", engine="openpyxl") as writer:
        kpi_df.to_excel(writer, sheet_name="KPI_Summary", index=False)
        df.to_excel(writer, sheet_name="Cleaned_Data", index=False)
        category_summary.to_excel(writer, sheet_name="Category_Analysis", index=False)

# -----------------------------
# Main Execution
# -----------------------------
def main():
    df = load_data()
    df = clean_data(df)
    kpi_df = calculate_kpis(df)
    category_summary = analyze_categories(df)
    plot_category_chart(category_summary)
    export_excel(df, kpi_df, category_summary)
    print("✅ Business performance report generated successfully.")

if __name__ == "__main__":
    main()
