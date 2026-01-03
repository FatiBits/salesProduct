import os
import pandas as pd

INPUT_FOLDER = r"D:\\Proje\\prj2\\data"   # مسیر واقعی خودت
OUTPUT_FOLDER = r"D:\\Proje\\prj2"
OUTPUT_FILE = "final_report.xlsx"

REQUIRED_COLUMNS = ["date", "product", "price", "quantity", "seller"]


def clean_dataframe(df):
    df.columns = (
        df.columns
        .str.strip()
        .str.lower()
        .str.replace(" ", "_")
    )

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        return pd.DataFrame()

    df["date"] = pd.to_datetime(df["date"], errors="coerce")

    df["price"] = (
        df["price"]
        .astype(str)
        .str.replace(",", "", regex=False)
        .str.strip()
    )
    df["price"] = pd.to_numeric(df["price"], errors="coerce")

    df["quantity"] = pd.to_numeric(df["quantity"], errors="coerce")

    df["seller"] = (
        df["seller"]
        .astype(str)
        .str.strip()
        .str.title()
    )

    df["product"] = (
        df["product"]
        .astype(str)
        .str.replace('"', '', regex=False)
        .str.strip()
        .str.title()
    )

    df = df.dropna(subset=REQUIRED_COLUMNS)

    return df


def read_and_clean_all_files(folder_path):
    dfs = []
    excel_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx")]

    for file in excel_files:
        path = os.path.join(folder_path, file)
        df = pd.read_excel(path)
        df_clean = clean_dataframe(df)
        if not df_clean.empty:
            dfs.append(df_clean)

    return dfs


def merge_dataframes(dfs):
    df = pd.concat(dfs, ignore_index=True)
    return df.drop_duplicates()


def calculate_metrics(df):
    df["total_sales"] = df["price"] * df["quantity"]

    kpi_df = pd.DataFrame({
        "Metric": ["Total Revenue", "Total Orders", "Average Order Value"],
        "Value": [
            df["total_sales"].sum(),
            len(df),
            df["total_sales"].mean()
        ]
    })

    sales_by_product = (
        df.groupby("product")["total_sales"]
        .sum()
        .sort_values(ascending=False)
        .reset_index()
    )

    sales_by_seller = (
        df.groupby("seller")["total_sales"]
        .sum()
        .sort_values(ascending=False)
        .reset_index()
    )

    return df, kpi_df, sales_by_product, sales_by_seller


def export_to_excel(df, kpi, by_product, by_seller):
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    output_path = os.path.join(OUTPUT_FOLDER, OUTPUT_FILE)

    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Clean_Data", index=False)
        kpi.to_excel(writer, sheet_name="KPI", index=False)
        by_product.to_excel(writer, sheet_name="Sales_By_Product", index=False)
        by_seller.to_excel(writer, sheet_name="Sales_By_Seller", index=False)

    print(f"📁 Report saved to: {output_path}")


if __name__ == "__main__":
    dfs = read_and_clean_all_files(INPUT_FOLDER)
    final_df = merge_dataframes(dfs)

    final_df, kpi_df, sales_by_product, sales_by_seller = calculate_metrics(final_df)

    export_to_excel(final_df, kpi_df, sales_by_product, sales_by_seller)


