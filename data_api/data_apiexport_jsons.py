import json
import os
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parents[1]
EXCEL_PATH = BASE_DIR / "Cubo_Semanal_Compactado.xlsx"
OUT_DIR = BASE_DIR / "data_api" / "out"

OUT_DIR.mkdir(parents=True, exist_ok=True)


def norm_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def find_col(cols, options):
    cols_map = {str(c).strip().lower(): c for c in cols}
    for opt in options:
        if opt.lower() in cols_map:
            return cols_map[opt.lower()]
    for opt in options:
        for k, v in cols_map.items():
            if opt.lower() in k or k in opt.lower():
                return v
    return None


def parse_month_key(v):
    if pd.isna(v):
        return None
    if isinstance(v, pd.Timestamp):
        return f"{v.year:04d}-{v.month:02d}"
    s = str(v).strip()
    try:
        dt = pd.to_datetime(s, errors="raise")
        return f"{dt.year:04d}-{dt.month:02d}"
    except Exception:
        pass

    s2 = s.replace("/", "-").replace("_", "-").replace(" ", "")
    if len(s2) == 7 and s2[4] == "-":
        y = s2[:4]
        m = s2[5:]
        if y.isdigit() and m.isdigit():
            return f"{int(y):04d}-{int(m):02d}"
    if len(s2) == 6 and s2.isdigit():
        return f"{int(s2[:4]):04d}-{int(s2[4:]):02d}"
    return None


def safe_write(name, obj):
    out_path = OUT_DIR / name
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)


def main():
    xls = pd.ExcelFile(EXCEL_PATH)
    sheets = xls.sheet_names

    sales_sheet = None
    for s in ["Cubo_Sales_DetailHallazgos", "Cubo_Sales", "Cubo_Sales_Fact"]:
        if s in sheets:
            sales_sheet = s
            break

    tx_sheet = None
    for s in ["Cubo_TX_Fact", "Cubo_TX", "Cubo_TX_Month_Fact"]:
        if s in sheets:
            tx_sheet = s
            break

    if not sales_sheet:
        raise ValueError("No encontré hoja de ventas")

    sales = norm_cols(pd.read_excel(EXCEL_PATH, sheet_name=sales_sheet))
    tx = norm_cols(pd.read_excel(EXCEL_PATH, sheet_name=tx_sheet)) if tx_sheet else None

    s_cols = list(sales.columns)
    t_cols = list(tx.columns) if tx is not None else []

    s_store = find_col(s_cols, ["Store", "Store Name", "Tienda"])
    s_period = find_col(s_cols, ["YearMonth", "YM", "Month", "Mes", "Fecha", "Date", "CalendarMonth", "CalMonth"])
    s_dept = find_col(s_cols, ["Department", "Departamento"])
    s_cat = find_col(s_cols, ["Category", "Categoria", "Categoría"])
    s_brand = find_col(s_cols, ["Brand", "Marca"])
    s_sales = find_col(s_cols, ["Sales", "Total Sales", "Venta", "Ventas"])
    s_margin = find_col(s_cols, ["Gross Margin", "Total Gross Margin", "Margen", "GM"])

    if not s_period or not s_sales:
        raise ValueError(f"No pude identificar columnas base en ventas: {s_cols}")

    sales["_month_key"] = sales[s_period].apply(parse_month_key)
    sales = sales[sales["_month_key"].notna()].copy()

    sales["_store"] = sales[s_store].astype(str).str.strip().str.upper() if s_store else "TOTAL"
    sales["_department"] = sales[s_dept].astype(str).str.strip() if s_dept else ""
    sales["_category"] = sales[s_cat].astype(str).str.strip() if s_cat else ""
    sales["_brand"] = sales[s_brand].astype(str).str.strip() if s_brand else ""
    sales["_sales"] = pd.to_numeric(sales[s_sales], errors="coerce").fillna(0)
    sales["_margin"] = pd.to_numeric(sales[s_margin], errors="coerce").fillna(0) if s_margin else 0

    monthly_category_sales = (
        sales.groupby(["_month_key", "_department", "_category"], dropna=False)["_sales"]
        .sum()
        .reset_index()
        .sort_values(["_month_key", "_department", "_sales"], ascending=[True, True, False])
    )
    monthly_category_sales["_category"] = monthly_category_sales["_category"].replace("", "Sin categoría")

    monthly_sales = (
        sales.groupby(["_month_key"], dropna=False)[["_sales", "_margin"]]
        .sum()
        .reset_index()
        .sort_values("_month_key")
    )

    tx_monthly = None
    if tx is not None:
        t_period = find_col(t_cols, ["YearMonth", "YM", "Month", "Mes", "Fecha", "Date", "CalendarMonth", "CalMonth"])
        t_tx = find_col(t_cols, ["Transactions", "TX", "Transacciones"])
        if t_period and t_tx:
            tx["_month_key"] = tx[t_period].apply(parse_month_key)
            tx = tx[tx["_month_key"].notna()].copy()
            tx["_tx"] = pd.to_numeric(tx[t_tx], errors="coerce").fillna(0)

            tx_monthly = (
                tx.groupby(["_month_key"], dropna=False)["_tx"]
                .sum()
                .reset_index()
                .sort_values("_month_key")
            )

    monthly_compare = monthly_sales.copy()
    if tx_monthly is not None:
        monthly_compare = monthly_compare.merge(tx_monthly, on="_month_key", how="left")
    else:
        monthly_compare["_tx"] = None

    safe_write("meta.json", {
        "sales_sheet": sales_sheet,
        "tx_sheet": tx_sheet,
        "months": sorted(monthly_sales["_month_key"].dropna().unique().tolist()),
        "rows_sales": int(len(sales)),
    })

    safe_write("monthly_category_sales.json", monthly_category_sales.to_dict(orient="records"))
    safe_write("monthly_compare.json", monthly_compare.to_dict(orient="records"))

    print("JSON exportados en:", OUT_DIR)


if __name__ == "__main__":
    main()