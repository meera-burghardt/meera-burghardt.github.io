# Load the ONS regional GDP workbooks and save tidy "long" CSVs
# Long format = one row per area per year, which is what the charts (and pandas) like best.

import pandas as pd
from pathlib import Path

DATA_DIR = Path(__file__).parent       # the folder this script lives in (CC4)
OUT_DIR = DATA_DIR / "data"            # CSVs are saved in CC4/data
OUT_DIR.mkdir(exist_ok=True)

FILES = {
    # short name: (file name, columns that identify the area)
    "itl": ("regionalgrossdomesticproductgdpbyallinternationalterritoriallevelitlregions.xlsx",
            ["ITL", "ITL code", "Region name"]),
    "la":  ("regionalgrossdomesticproductgdplocalauthorities.xlsx",
            ["ITL1 Region", "LA code", "LA name"]),
    "cer": ("regionalgrossdomesticproductgdpcityandenterpriseregions.xlsx",
            ["Area type", "Geo code", "Area name"]),
}

# The tables we need (see the Contents sheet for the full list)
TABLES = {
    "table5":  "gdp_current_prices_gbp_m",
    "table6":  "population",
    "table11": "real_gdp_per_head_2023_gbp",
}


def load_table(path, sheet, id_cols):
    """Read one table and return it in long format with numeric values."""
    df = pd.read_excel(path, sheet_name=sheet, header=1)     # row 2 of the sheet holds the headers
    df.columns = [str(c).strip() for c in df.columns]        # year headers can come in as text
    year_cols = [c for c in df.columns if c.isdigit()]       # 1998 ... 2024
    long = df.melt(id_vars=id_cols, value_vars=year_cols, var_name="year", value_name="value")
    long["year"] = long["year"].astype(int)
    long["value"] = pd.to_numeric(long["value"], errors="coerce")   # [z] and [u] markers become NaN
    return long


for key, (fname, id_cols) in FILES.items():
    for sheet_key, label in TABLES.items():
        sheet = sheet_key.replace("table", "Table ")          # "table11" -> "Table 11"
        long = load_table(DATA_DIR / fname, sheet, id_cols)
        out = OUT_DIR / f"{key}_{label}.csv"
        long.to_csv(out, index=False)
        print(f"saved {out}  ({len(long):,} rows, {long['value'].isna().sum()} missing values)")
