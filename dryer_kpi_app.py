import pandas as pd
import numpy as np
from pathlib import Path
import xlsxwriter

# CONFIG remains unchanged (lines 7-19)
CONFIG = {
    "energy_file": r"E:\Lindner\Python\Energieverbrauch Trockner 1, Stundenweise - Januar - September 2025.xlsx",
    "energy_sheet": 0,
    "wagon_file": r"E:\Lindner\Python\Hordenwagenverfolgung_Stand 2025_10_12.xlsm",
    "wagon_sheet": "Hordenwagenverfolgung",
    "wagon_header_row": 6,
    "gas_to_kwh": 11.5,
    "takt_minutes": 65,
    "zones_seq": ["Z1","Z2","Z3","Z4","Z5"],
    "product_filter": ["L36"],     # or None
    "month_filter": None,          # e.g. 8 for August only
    "output_file": r"E:\Lindner\Python\Dryer_KPI_Monthly_Results.xlsx"
}

# ---------- Helper Functions (Unchanged - lines 21-197) ----------
def parse_energy(df):
# ... (function body unchanged) ...
    df = df.copy()
    df["Zeitstempel"] = pd.to_datetime(df["Zeitstempel"])
    df["Month"] = df["Zeitstempel"].dt.month
    for z in ["Zone 2", "Zone 3", "Zone 4", "Zone 5"]:
        col = f"Gasmenge, {z} [m³]"
        if col in df.columns:
            df[f"E_{z}_kWh"] = df[col] * CONFIG["gas_to_kwh"]
    if "Energieverbrauch, elektr. [kWh]" in df.columns:
        df["E_el_kWh"] = df["Energieverbrauch, elektr. [kWh]"]
    df["E_start"] = df["Zeitstempel"]
    df["E_end"] = df["Zeitstempel"] + pd.Timedelta(hours=1)
    return df

def parse_duration_series(s: pd.Series) -> pd.Series:
# ... (function body unchanged) ...
    """
    Convert the free-text “Zeit in Zx” column (e.g. “12:34”, “5 h 30 min”, …)
    into a pandas Timedelta Series.
    Return NaT where parsing fails.
    """
    # Replace common abbreviations, then let pandas do the heavy lifting
    s = s.astype(str).str.strip()
    s = s.replace({
        r'\bh\b': 'h', r'\bmin\b': 'm', r'\bst\b': 's',
        r'^\s*$': np.nan, r'^-$': np.nan
    }, regex=True)

    # pandas can parse most ISO-like strings directly
    return pd.to_timedelta(s, errors='coerce')


def parse_wagon(df):
# ... (function body unchanged) ...
    df = df.copy()
    # ... (all wagon parsing logic remains unchanged) ...

    # ------------------------------------------------------------------ #
    # 9. Helper column for monthly aggregation
    # ------------------------------------------------------------------ #
    df["Month"] = df["t0"].dt.month

    return df


def build_intervals(row):
# ... (function body unchanged) ...
    intervals = []
    prev_end = None
    for z in CONFIG["zones_seq"]:
        zin = row.get(f"{z}_in", pd.NaT)
        if pd.isna(zin):
            zin = prev_end if prev_end is not None else row["t0"]
        zdur = row.get(f"{z}_dur", pd.NaT)
        zout = zin + zdur if pd.notna(zin) and pd.notna(zdur) else pd.NaT
        if pd.notna(zin) and pd.notna(zout) and zout > zin:
            intervals.append((z, zin, zout))
            prev_end = zout
    return intervals

def explode_intervals(df):
# ... (function body unchanged) ...
    rows = []
    for _, r in df.iterrows():
        ivals = build_intervals(r)
        for z, a, b in ivals:
            rows.append({
                "WG_Nr": r["WG_Nr"], "Produkt": r["Produkt"], "Stärke": r["Stärke"],
                "m3": r["m3"], "Zone": z, "P_start": a, "P_end": b, "Month": r["Month"]
            })
    return pd.DataFrame(rows)

def overlap_hours(a_start, a_end, b_start, b_end):
# ... (function body unchanged) ...
    latest = max(a_start, b_start)
    earliest = min(a_end, b_end)
    dh = (earliest - latest).total_seconds() / 3600.0
    return max(0.0, dh)

def allocate_energy(e, ivals):
# ... (function body unchanged) ...
    out = []
    zone_cols = [("Z2","E_Zone 2_kWh"),("Z3","E_Zone 3_kWh"),("Z4","E_Zone 4_kWh"),("Z5","E_Zone 5_kWh")]
    for _, er in e.iterrows():
        E_start, E_end, month = er["E_start"], er["E_end"], er["Month"]
        for zlabel, zcol in zone_cols:
            if zcol not in e.columns: continue
            E_hour = er[zcol]
            if pd.isna(E_hour) or E_hour == 0: continue
            sub = ivals[(ivals["Zone"]==zlabel) & (ivals["P_end"]>E_start) & (ivals["P_start"]<E_end)]
            for _, pr in sub.iterrows():
                ovh = overlap_hours(E_start, E_end, pr["P_start"], pr["P_end"])
                if ovh > 0:
                    out.append({
                        "Month": month, "Zone": zlabel, "Produkt": pr["Produkt"],
                        "Energy_share_kWh": E_hour * ovh, "Overlap_h": ovh,
                        "m3": pr["m3"]
                    })
    return pd.DataFrame(out)

# ---------- Main (Modified with Error Handling) ----------
def main():
    e_raw = pd.read_excel(CONFIG["energy_file"], sheet_name=CONFIG["energy_sheet"])
    e = parse_energy(e_raw)
    w_raw = pd.read_excel(CONFIG["wagon_file"], sheet_name=CONFIG["wagon_sheet"], header=CONFIG["wagon_header_row"])
    w = parse_wagon(w_raw)
    
    # --- 1. Apply Filters ---
    if CONFIG["product_filter"]:
        w = w[w["Produkt"].astype(str).isin(CONFIG["product_filter"])]
    if CONFIG["month_filter"]:
        e = e[e["Month"] == CONFIG["month_filter"]]
        w = w[w["Month"] == CONFIG["month_filter"]]
    
    # --- 2. Check for empty DataFrames after filtering ---
    if w.empty:
        # Raise an exception with a descriptive error message
        msg = "Wagon data is empty after applying filters or failed to load/parse. Check 'wagon_sheet', 'wagon_header_row', 'product_filter', and 'month_filter' in the configuration or the Streamlit inputs."
        print(f"ERROR: {msg}")
        raise ValueError(msg) 
    
    # --- 3. Run core analysis steps ---
    ivals = explode_intervals(w)
    
    if ivals.empty:
        # This catches cases where the wagon data is present but the timestamps 
        # (Z_in, Z_dur) are all invalid, preventing interval creation.
        msg = "Interval data is empty. Wagon data loaded successfully, but failed to create valid Zone intervals. Check timestamp columns and time duration columns in the Hordenwagen file."
        print(f"ERROR: {msg}")
        raise ValueError(msg)
    
    alloc = allocate_energy(e, ivals)

    # Monthly summary
    summary = alloc.groupby(["Month","Produkt","Zone"], as_index=False).agg(
        Energy_kWh=("Energy_share_kWh","sum"),
        Volume_m3=("m3","sum")
    )
    summary["kWh_per_m3"] = summary["Energy_kWh"] / summary["Volume_m3"].replace(0, np.nan)

    # Yearly summary
    yearly = summary.groupby(["Produkt","Zone"], as_index=False).agg(
        Energy_kWh=("Energy_kWh","sum"),
        Volume_m3=("Volume_m3","sum")
    )
    yearly["kWh_per_m3"] = yearly["Energy_kWh"] / yearly["Volume_m3"].replace(0, np.nan)

    # ---------- Excel Output (Unchanged - lines 270-302) ----------
    with pd.ExcelWriter(CONFIG["output_file"], engine="xlsxwriter") as writer:
        e.to_excel(writer, sheet_name="Energy_Hourly_Parsed", index=False)
        w.to_excel(writer, sheet_name="Wagons_Parsed", index=False)
        ivals.to_excel(writer, sheet_name="Intervals_By_Zone", index=False)
        alloc.to_excel(writer, sheet_name="Energy_Allocated", index=False)
        summary.to_excel(writer, sheet_name="Summary_By_Month_Zone", index=False)
        yearly.to_excel(writer, sheet_name="Yearly_Summary", index=False)

        wb = writer.book
        fmt_head = wb.add_format({'bold': True, 'bg_color': '#C6E0B4', 'border':1})
        fmt_num = wb.add_format({'num_format': '#,##0.00', 'border':1})
        for s in ["Summary_By_Month_Zone","Yearly_Summary"]:
            ws = writer.sheets[s]
            ws.set_row(0, 18, fmt_head)
            ws.set_column("A:F", 18, fmt_num)

        # Charts
        ws_sum = writer.sheets["Summary_By_Month_Zone"]
        chart = wb.add_chart({'type':'column'})
        for z in ["Z2","Z3","Z4","Z5"]:
            data = summary[summary["Zone"]==z]
            if data.empty: continue
            start = summary.index[summary["Zone"]==z].min() + 1
            chart.add_series({
                'name': z,
                'categories': ['Summary_By_Month_Zone', start, 0, start+len(data)-1, 0],
                'values': ['Summary_By_Month_Zone', start, 5, start+len(data)-1, 5]
            })
        chart.set_title({'name': 'Monthly KPI (kWh/m³)'})
        chart.set_x_axis({'name': 'Month'})
        chart.set_y_axis({'name': 'kWh/m³'})
        chart.set_style(10)
        ws_sum.insert_chart('H2', chart)

        ws_year = writer.sheets["Yearly_Summary"]
        chart2 = wb.add_chart({'type': 'column'})
        chart2.add_series({
            'name': 'Yearly KPI (kWh/m³)',
            'categories': ['Yearly_Summary', 1, 1, len(yearly), 1],
            'values': ['Yearly_Summary', 1, 3, len(yearly), 3]
        })
        chart2.set_title({'name': 'Yearly KPI per Zone'})
        chart2.set_style(10)
        ws_year.insert_chart('H2', chart2)

if __name__ == "__main__":
    main()
