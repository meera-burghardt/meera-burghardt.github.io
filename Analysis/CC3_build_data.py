# CC3 data builder: cleans the ONS files and writes the two CSVs the charts read.
#
# Chart 1 (nations_gdp.csv):
#     Real GDP per head, 1998 = 100, for Scotland, Wales, Northern Ireland and England excluding London.
#     Source: ONS regional GDP by ITL region (xlsx), Table 10 (real GDP, 2023 prices) and Table 6 (population).
#
# Chart 2 (productivity_dumbbell.csv):
#     Productivity (GVA per hour worked, current prices) in 2004 vs 2023, for areas with and without a mayor.
#     Source: ONS subregional productivity by local authority (xls), sheets A3 (GVA per hour) and Productivity Hours.
#
# Folder layout this script expects:
#     analysis/cc3_build_data.py      <- this script
#     analysis/raw/                   <- the ONS .xlsx and .xls files
#     portfolio/                      <- the website; the CSVs are written here
#
# Set up once:   python3 -m pip install pandas openpyxl xlrd
# Run:           python3 analysis/cc3_build_data.py        (works from any folder)

from pathlib import Path
import json
import pandas as pd

# ======================================================================================
# 0. SETTINGS (change these if you want to try a different choice)
# ======================================================================================
HERE = Path(__file__).resolve().parent      # the analysis folder
RAW = HERE / "raw"                          # where the ONS files live
OUT = HERE.parent / "portfolio"             # where the charts look for their CSVs

# Chart 1
BASE_YEAR = 1998                            # index = 100 in this year (last year before devolution in 1999)

# Chart 2
START, END = 2004, 2023                     # the two years shown on the dumbbell
RANK_YEAR = 2013                            # year used to rank candidate comparison areas by size (hours worked)
N_COMPARATORS = 6                           # how many individual no-mayor areas to show


def find_file(pattern):
    """Find a raw file by name pattern (works even if the browser added '(1)' to the name)."""
    matches = sorted(RAW.glob(pattern))
    if not matches:
        raise SystemExit(f"Could not find '{pattern}' in {RAW}. Put the ONS file in the raw folder.")
    return matches[0]


# ======================================================================================
# 1. CHART 1: real GDP per head for the nations, indexed to BASE_YEAR = 100
# ======================================================================================
def build_nations():
    itl_file = find_file("regionalgrossdomesticproductgdpbyallinternationalterritoriallevelitlregions*.xlsx")

    def read_table(sheet):
        df = pd.read_excel(itl_file, sheet_name=sheet, header=1)    # row 2 holds the column names
        df.columns = [str(c).strip() for c in df.columns]           # year headers can arrive as text
        years = [c for c in df.columns if c.isdigit()]
        for c in years:
            df[c] = pd.to_numeric(df[c], errors="coerce")           # [z] and [u] markers become NaN
        return df.set_index("ITL code"), years

    real_gdp, years = read_table("Table 10")    # real GDP, £ million, 2023 prices
    population, _ = read_table("Table 6")       # population

    # Which ITL1 regions make up each line (TLI = London, left out of England on purpose)
    groups = {
        "Scotland": ["TLM"],
        "Wales": ["TLL"],
        "Northern Ireland": ["TLN"],
        "England excl. London": ["TLC", "TLD", "TLE", "TLF", "TLG", "TLH", "TLJ", "TLK"],
    }

    rows = []
    for area, codes in groups.items():
        # Real GDP per head = real GDP / population, summed over the regions in the group.
        # (£ million per person x 1,000,000 = £ per person.)
        per_head = real_gdp.loc[codes, years].sum() / population.loc[codes, years].sum() * 1_000_000
        for year in years:
            rows.append({
                "year": int(year),
                "area": area,
                "gdp_per_head_gbp": round(per_head[year]),
                "index_1998": round(per_head[year] / per_head[str(BASE_YEAR)] * 100, 1),
            })

    nations = pd.DataFrame(rows)
    path = OUT / "nations_gdp.csv"
    nations.to_csv(path, index=False)
    print(f"Chart 1: saved {path} ({len(nations)} rows)")
    return nations


# ======================================================================================
# 2. CHART 2: productivity (GVA per hour) in START vs END, mayoral vs no-mayor areas
# ======================================================================================
def build_productivity():
    prod_file = find_file("labourproductivitylad1*.xls")

    def load(sheet):
        df = pd.read_excel(prod_file, sheet_name=sheet, header=4)           # column names sit on row 5
        df = df[df["LAD_Name"].notna()]                                      # drop blank and footnote rows
        df = df[df["LAD_Code"].astype(str).str.match(r"^[A-Z]\d")]           # keep real local authority codes
        return df.set_index("LAD_Code")

    pounds = load("A3")                  # GVA per hour worked, current prices (£)
    hours = load("Productivity Hours")   # total hours worked, used as weights
    names = pounds["LAD_Name"].str.strip()

    def codes_for(local_authorities):
        """Look up LAD codes from names (and warn if a name is not found)."""
        found = names[names.isin(local_authorities)].index.tolist()
        if len(found) != len(local_authorities):
            print("  Check these names:", set(local_authorities) - set(names[found]))
        return found

    def productivity(codes, year):
        """Hours-weighted average GVA per hour for a group of local authorities."""
        p = pounds.loc[codes, f"Pounds_{year}"].astype(float)
        h = hours.loc[codes, f"Hours_{year}"].astype(float)
        return (p * h).sum() / h.sum()

    # --- Groups --------------------------------------------------------------------
    # Areas with a mayor (the "devolved" English areas)
    combined_authorities = {
        "Greater Manchester": ["Bolton", "Bury", "Manchester", "Oldham", "Rochdale", "Salford",
                               "Stockport", "Tameside", "Trafford", "Wigan"],
        "West Midlands": ["Birmingham", "Coventry", "Dudley", "Sandwell", "Solihull", "Walsall", "Wolverhampton"],
        "Liverpool City Region": ["Halton", "Knowsley", "Liverpool", "St. Helens", "Sefton", "Wirral"],
        "West Yorkshire": ["Bradford", "Calderdale", "Kirklees", "Leeds", "Wakefield"],
        "South Yorkshire": ["Barnsley", "Doncaster", "Rotherham", "Sheffield"],
        "Tees Valley": ["Hartlepool", "Middlesbrough", "Redcar and Cleveland", "Stockton-on-Tees", "Darlington"],
    }
    # Other areas that had a mayor by 2023 (kept out of the no-mayor group)
    other_mayoral = ["Cambridge", "East Cambridgeshire", "Fenland", "Huntingdonshire", "South Cambridgeshire",
                     "Peterborough", "Bath and North East Somerset", "Bristol, City of", "South Gloucestershire",
                     "Newcastle upon Tyne", "North Tyneside", "Northumberland"]

    mayoral = {k: codes_for(v) for k, v in combined_authorities.items()}
    all_mayoral = sum(mayoral.values(), []) + codes_for(other_mayoral)

    english = [c for c in pounds.index if c.startswith("E") and not c.startswith("E09")]   # E09 = London
    no_mayor = [c for c in english if c not in all_mayoral]

    # Comparison areas: the largest no-mayor areas by hours worked (a rule, not hand-picking)
    largest = hours.loc[no_mayor, f"Hours_{RANK_YEAR}"].astype(float).sort_values(ascending=False)
    comparators = largest.head(N_COMPARATORS).index.tolist()

    # --- One row per area ----------------------------------------------------------
    rows = []

    def add(label, codes, group, pooled):
        start, end = productivity(codes, START), productivity(codes, END)
        rows.append({"area": label, "group": group, "pooled": pooled,
                     "prod_start": round(start, 2), "prod_end": round(end, 2),
                     "change_gbp": round(end - start, 2),
                     "growth_pct": round((end / start - 1) * 100, 1)})

    for label, codes in mayoral.items():
        add(label, codes, "Mayoral combined authority", 0)
    add("Six mayoral areas (pooled)", sum(mayoral.values(), []), "Mayoral combined authority", 1)
    for code in comparators:
        add(names[code], [code], "No mayor", 0)
    add("Rest of England, no mayor (pooled)", no_mayor, "No mayor", 1)

    df = pd.DataFrame(rows)

    # --- Row order for the chart: no-mayor group on top, mayoral group below ----------
    # Within each group the pooled row comes first, then areas sorted by growth.
    df["group_rank"] = df["group"].map({"No mayor": 0, "Mayoral combined authority": 1})
    df = df.sort_values(["group_rank", "pooled", "growth_pct"], ascending=[True, False, False])
    df["sort_order"] = range(1, len(df) + 1)
    df = df.drop(columns="group_rank")

    path = OUT / "productivity_dumbbell.csv"
    df.to_csv(path, index=False)
    print(f"Chart 2: saved {path} ({len(df)} rows)")
    return df


# ======================================================================================
# 3. RUN
# ======================================================================================
if __name__ == "__main__":
    OUT.mkdir(exist_ok=True)
    nations = build_nations()
    productivity_table = build_productivity()

    print("\nChart 1 check (index, 1998 / 2023 / 2024):")
    print(nations.pivot(index="year", columns="area", values="index_1998").loc[[1998, 2023, 2024]])
    print("\nChart 2 check:")
    print(productivity_table[["area", "group", "prod_start", "prod_end", "growth_pct"]].to_string(index=False))

    # The chart's row order is written into wk2_chart2.json (the "sort" lists).
    # If you change the comparison areas or rules above, paste this new list over those "sort" lists.
    print("\nRow order for wk2_chart2.json:")
    print(json.dumps(productivity_table["area"].tolist()))