"""
Lindner Dryer - Simplified KPI Analysis
Standalone version with all functions embedded
"""

import streamlit as st
import pandas as pd
import numpy as np
import tempfile
import plotly.express as px
import os

st.set_page_config(page_title="KPI Analysis", page_icon="📊", layout="wide")

# ===== EMBEDDED CONFIGURATION =====
CONFIG = {
    "energy_sheet": 0,
    "wagon_sheet": "Hordenwagenverfolgung",
    "wagon_header_row": 6,
    "gas_to_kwh": 11.5,
}

ZONE_MAPPING = {
    "Z2": "Zone 2",
    "Z3": "Zone 3",
    "Z4": "Zone 4",
    "Z5": "Zone 5"
}

# ===== EMBEDDED FUNCTIONS =====

def parse_energy_simple(df):
    """Parse energy data - simplified"""
    df = df.copy()
    
    # Parse timestamp
    df["Zeitstempel"] = pd.to_datetime(df["Zeitstempel"], errors='coerce')
    df["Month"] = df["Zeitstempel"].dt.month
    
    # Convert gas to kWh
    for zone_key, zone_name in ZONE_MAPPING.items():
        gas_col = f"Gasmenge, {zone_name} [m³]"
        if gas_col in df.columns:
            df[f"E_{zone_name}_kWh"] = df[gas_col] * CONFIG["gas_to_kwh"]
    
    return df[df["Zeitstempel"].notna()].copy()

def parse_wagon_simple(df):
    """Parse wagon data - simplified"""
    df = df.copy()
    
    # Clean column names
    df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]
    
    # Find wagon number column
    for col in df.columns:
        if col.startswith("WG-"):
            df = df.rename(columns={col: "WG_Nr"})
            break
    
    # Calculate volume
    if "m³" in df.columns:
        df["m3"] = pd.to_numeric(df["m³"], errors='coerce')
    else:
        if "Stärke" in df.columns:
            staerke = pd.to_numeric(df["Stärke"], errors='coerce').fillna(36)
            df["m3"] = 0.605 * 0.605 * (staerke + 7) / 1000
        else:
            df["m3"] = 0.605 * 0.605 * (36 + 7) / 1000
    
    # Add month if timestamp exists
    if "Pressen-Datum" in df.columns or "Press-Zeit" in df.columns:
        try:
            df["Month"] = pd.to_datetime(df.get("Pressen-Datum", ""), errors='coerce').dt.month
        except:
            df["Month"] = 1
    else:
        df["Month"] = 1
    
    return df

def simple_kpi_analysis(energy_df, wagon_df, products_filter=None, month_filter=None):
    """Simple KPI calculation"""
    
    # Parse data
    energy = parse_energy_simple(energy_df)
    wagons = parse_wagon_simple(wagon_df)
    
    # Apply filters
    if products_filter and "Produkt" in wagons.columns:
        wagons = wagons[wagons["Produkt"].isin(products_filter)]
    
    if month_filter and "Month" in wagons.columns:
        wagons = wagons[wagons["Month"] == month_filter]
    
    if month_filter and "Month" in energy.columns:
        energy = energy[energy["Month"] == month_filter]
    
    # Calculate KPIs by product and zone
    results = []
    
    if "Produkt" not in wagons.columns:
        return None
    
    for product in wagons["Produkt"].dropna().unique():
        product_wagons = wagons[wagons["Produkt"] == product]
        total_volume = product_wagons["m3"].sum()
        wagon_count = len(product_wagons)
        
        for zone_key, zone_name in ZONE_MAPPING.items():
            energy_col = f"E_{zone_name}_kWh"
            
            if energy_col in energy.columns:
                total_energy = energy[energy_col].sum()
                
                # Simple allocation: proportional to volume
                total_wagon_volume = wagons["m3"].sum()
                if total_wagon_volume > 0:
                    product_energy = total_energy * (total_volume / total_wagon_volume)
                else:
                    product_energy = 0
                
                results.append({
                    'Produkt': product,
                    'Zone': zone_key,
                    'Energy_kWh': product_energy,
                    'Volume_m3': total_volume,
                    'Wagons': wagon_count,
                    'kWh_per_m3': product_energy / total_volume if total_volume > 0 else 0
                })
    
    return pd.DataFrame(results)

# ===== UI =====

st.title("📊 Lindner Dryer - KPI Analysis")
st.info("Upload your files to analyze energy efficiency")

# Sidebar
with st.sidebar:
    st.image("https://www.karrieretag.org/wp-content/uploads/2023/10/lindner-logo-1.png", 
             use_container_width=True)
    st.markdown("---")
    
    st.subheader("📁 Upload Files")
    energy_file = st.file_uploader("Energy File (.xlsx)", type=["xlsx"], key="energy")
    wagon_file = st.file_uploader("Wagon File (.xlsm, .xlsx)", type=["xlsm", "xlsx"], key="wagon")
    
    st.markdown("---")
    st.subheader("⚙️ Filters")
    
    products_list = ["L28", "L30", "L32", "L34", "L36", "L38", "L40", "L44", "N40", "N44", "U36"]
    
    select_all = st.checkbox("Select All Products", value=True, key="select_all")
    
    if select_all:
        selected_products = products_list
        st.multiselect("Products:", products_list, default=products_list, disabled=True, key="products_disabled")
    else:
        selected_products = st.multiselect("Products:", products_list, default=["L36"], key="products_manual")
    
    st.info(f"Selected: {len(selected_products)} products")
    
    month = st.number_input("Month (0 = all):", 0, 12, 0, key="month")
    
    st.markdown("---")
    analyze_btn = st.button("▶️ Run Analysis", use_container_width=True, type="primary")

# Main analysis
if analyze_btn:
    if not energy_file or not wagon_file:
        st.error("⚠️ Please upload both files")
    else:
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_e, \
             tempfile.NamedTemporaryFile(suffix=".xlsm", delete=False) as tmp_w:
            
            # Save uploads to temp files
            tmp_e.write(energy_file.read())
            tmp_w.write(wagon_file.read())
            tmp_e.flush()
            tmp_w.flush()
            
            try:
                # Progress
                progress = st.progress(0)
                status = st.empty()
                
                # Load files
                status.text("📊 Loading energy data...")
                progress.progress(25)
                energy_df = pd.read_excel(tmp_e.name, sheet_name=CONFIG["energy_sheet"])
                
                status.text("🚛 Loading wagon data...")
                progress.progress(50)
                wagon_df = pd.read_excel(
                    tmp_w.name,
                    sheet_name=CONFIG["wagon_sheet"],
                    header=CONFIG["wagon_header_row"]
                )
                
                # Analyze
                status.text("🔄 Calculating KPIs...")
                progress.progress(75)
                
                results_df = simple_kpi_analysis(
                    energy_df,
                    wagon_df,
                    selected_products if selected_products else None,
                    month if month != 0 else None
                )
                
                progress.progress(100)
                status.empty()
                progress.empty()
                
                if results_df is None or results_df.empty:
                    st.warning("⚠️ No data found. Please check your filters and files.")
                else:
                    st.success("✅ Analysis complete!")
                    
                    # Summary KPIs
                    st.markdown("### 📈 Key Performance Indicators")
                    
                    col1, col2, col3 = st.columns(3)
                    
                    total_energy = results_df['Energy_kWh'].sum()
                    total_volume = results_df['Volume_m3'].sum()
                    avg_kpi = results_df['kWh_per_m3'].mean()
                    
                    col1.metric("Total Energy", f"{total_energy:,.0f} kWh")
                    col2.metric("Total Volume", f"{total_volume:,.0f} m³")
                    col3.metric("Avg Efficiency", f"{avg_kpi:.2f} kWh/m³")
                    
                    # Charts
                    st.markdown("### 📊 Analysis Charts")
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        # Group by zone and product
                        fig1 = px.bar(
                            results_df,
                            x="Zone",
                            y="kWh_per_m3",
                            color="Produkt",
                            title="Energy Efficiency by Zone & Product",
                            barmode="group"
                        )
                        fig1.update_layout(height=400)
                        st.plotly_chart(fig1, use_container_width=True)
                    
                    with col2:
                        # Pie chart
                        fig2 = px.pie(
                            results_df,
                            values="Energy_kWh",
                            names="Produkt",
                            title="Energy Distribution by Product"
                        )
                        fig2.update_layout(height=400)
                        st.plotly_chart(fig2, use_container_width=True)
                    
                    # Data table
                    with st.expander("📋 View Detailed Data"):
                        st.dataframe(
                            results_df.style.format({
                                'Energy_kWh': '{:,.2f}',
                                'Volume_m3': '{:,.2f}',
                                'kWh_per_m3': '{:,.2f}',
                                'Wagons': '{:,.0f}'
                            }),
                            use_container_width=True
                        )
                    
                    # Export
                    st.markdown("### 📥 Export Results")
                    
                    # Excel export
                    output_file = "kpi_results.xlsx"
                    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                        results_df.to_excel(writer, sheet_name='KPI_Results', index=False)
                    
                    with open(output_file, 'rb') as f:
                        st.download_button(
                            "📥 Download Excel Report",
                            f.read(),
                            output_file,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    
                    # Cleanup
                    os.unlink(output_file)
                
            except Exception as e:
                st.error(f"❌ Error during analysis: {str(e)}")
                
                with st.expander("🔍 Error Details"):
                    st.exception(e)
            
            finally:
                # Cleanup temp files
                try:
                    os.unlink(tmp_e.name)
                    os.unlink(tmp_w.name)
                except:
                    pass

else:
    # Instructions
    st.markdown("""
    ## 🚀 How to Use
    
    1. **Upload Files** in the sidebar:
       - Energy consumption file (.xlsx)
       - Hordenwagen tracking file (.xlsm or .xlsx)
    
    2. **Select Products** to analyze (or "Select All")
    
    3. **Optional:** Filter by specific month (1-12) or 0 for all
    
    4. **Click "Run Analysis"**
    
    ---
    
    ## 📊 What You'll Get:
    
    - ✅ Total energy consumption
    - ✅ Energy efficiency (kWh/m³) per product and zone
    - ✅ Volume analysis
    - ✅ Visual charts and comparisons
    - ✅ Downloadable Excel report
    
    ---
    
    ### 📁 File Requirements:
    
    **Energy File:**
    - Hourly energy data
    - Columns: Zeitstempel, Gasmenge Zone 2-5
    
    **Wagon File:**
    - Production tracking data
    - Header starts at row 7
    - Columns: WG-Nr, Produkt, Stärke, etc.
    """)

st.markdown("---")
st.caption("🏭 Lindner Dryer KPI Analysis - Simplified Version")
