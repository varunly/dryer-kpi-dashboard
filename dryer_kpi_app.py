"""
Lindner Dryer - KPI Analysis Dashboard
Analyzes energy efficiency and creates reports
"""

import streamlit as st
import pandas as pd
import tempfile
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import os

# Import KPI calculation module
from dryer_kpi_monthly_final import (
    parse_energy, parse_wagon, explode_intervals, 
    allocate_energy, CONFIG
)

# ------------------ Page Configuration ------------------
st.set_page_config(
    page_title="Lindner Dryer - KPI Analysis",
    page_icon="📊",
    layout="wide"
)

# ------------------ Custom CSS ------------------
st.markdown("""
    <style>
    .main-title {
        font-size: 36px;
        color: #003366;
        font-weight: 700;
        text-align: center;
        margin-bottom: 20px;
    }
    
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 20px;
        border-radius: 15px;
        text-align: center;
        box-shadow: 0 4px 15px rgba(0,0,0,0.2);
        color: white;
    }
    
    .metric-card h3 {
        margin: 0;
        font-size: 16px;
        opacity: 0.9;
    }
    
    .metric-card h2 {
        margin: 10px 0 0 0;
        font-size: 32px;
        font-weight: 700;
    }
    </style>
""", unsafe_allow_html=True)

# ------------------ Helper Functions ------------------
def create_kpi_card(title, value, unit):
    """Create a styled KPI metric card"""
    if isinstance(value, (int, float)):
        formatted_value = f"{value:,.2f}"
    else:
        formatted_value = str(value)
    
    return f'''
    <div class="metric-card">
        <h3>{title}</h3>
        <h2>{formatted_value} {unit}</h2>
    </div>
    '''

def run_kpi_analysis(energy_path, wagon_path, products_filter, month_filter):
    """Run KPI analysis"""
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        # Parse energy data
        status_text.text("🔄 Parsing energy data...")
        progress_bar.progress(20)
        e_raw = pd.read_excel(energy_path, sheet_name=CONFIG["energy_sheet"])
        e = parse_energy(e_raw)
        
        # Parse wagon data
        status_text.text("🔄 Parsing wagon tracking data...")
        progress_bar.progress(40)
        w_raw = pd.read_excel(
            wagon_path, 
            sheet_name=CONFIG["wagon_sheet"], 
            header=CONFIG["wagon_header_row"]
        )
        w = parse_wagon(w_raw)
        
        # Apply filters
        status_text.text("🔄 Applying filters...")
        progress_bar.progress(60)
        if products_filter:
            w = w[w["Produkt"].astype(str).isin(products_filter)]
        if month_filter:
            e = e[e["Month"] == month_filter]
            w = w[w["Month"] == month_filter]
        
        # Process intervals
        status_text.text("🔄 Processing intervals...")
        progress_bar.progress(70)
        ivals = explode_intervals(w)
        
        # Allocate energy
        status_text.text("🔄 Allocating energy...")
        progress_bar.progress(85)
        alloc = allocate_energy(e, ivals)
        
        # Create summaries
        status_text.text("🔄 Creating summaries...")
        progress_bar.progress(95)
        
        summary = alloc.groupby(["Month", "Produkt", "Zone"], as_index=False).agg(
            Energy_kWh=("Energy_share_kWh", "sum"),
            Volume_m3=("m3", "sum")
        )
        summary["kWh_per_m3"] = summary["Energy_kWh"] / summary["Volume_m3"].replace(0, pd.NA)
        
        yearly = summary.groupby(["Produkt", "Zone"], as_index=False).agg(
            Energy_kWh=("Energy_kWh", "sum"),
            Volume_m3=("Volume_m3", "sum")
        )
        yearly["kWh_per_m3"] = yearly["Energy_kWh"] / yearly["Volume_m3"].replace(0, pd.NA)
        
        progress_bar.progress(100)
        status_text.text("✅ Analysis complete!")
        
        progress_bar.empty()
        status_text.empty()
        
        return {
            'summary': summary,
            'yearly': yearly,
            'success': True
        }
        
    except Exception as e:
        status_text.empty()
        progress_bar.empty()
        return {'success': False, 'error': str(e)}

# ------------------ Header ------------------
st.markdown('<div class="main-title">📊 Lindner Dryer - KPI Analysis</div>', 
            unsafe_allow_html=True)

st.info("Upload your energy and wagon files to analyze dryer efficiency")

# ------------------ Sidebar ------------------
with st.sidebar:
    st.image("https://www.karrieretag.org/wp-content/uploads/2023/10/lindner-logo-1.png", 
             use_column_width=True)
    st.markdown("---")
    
    st.subheader("📁 Upload Files")
    energy_file = st.file_uploader("Energy File (.xlsx)", type=["xlsx"])
    wagon_file = st.file_uploader("Hordenwagen File (.xlsm, .xlsx)", type=["xlsm", "xlsx"])
    
    st.markdown("---")
    st.subheader("⚙️ Filters")
    
    all_products = ["L28", "L30", "L32", "L34", "L36", "L38", "L40", "L44", "N40", "N44", "U36"]
    
    select_all = st.checkbox("Select All Products", value=True)
    
    if select_all:
        products = all_products
        st.multiselect("Products:", all_products, default=all_products, disabled=True, key="prod_disabled")
    else:
        products = st.multiselect("Products:", all_products, default=["L36"], key="prod_manual")
    
    st.info(f"Selected: {len(products)} products")
    
    month = st.number_input("Month (0 = all):", 0, 12, 0)
    
    st.markdown("---")
    run_button = st.button("▶️ Run Analysis", use_container_width=True)

# ------------------ Main Analysis ------------------
if run_button:
    if not energy_file or not wagon_file:
        st.error("⚠️ Please upload both files")
    else:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_e, \
             tempfile.NamedTemporaryFile(delete=False, suffix=".xlsm") as tmp_w:
            
            tmp_e.write(energy_file.read())
            tmp_w.write(wagon_file.read())
            tmp_e.flush()
            tmp_w.flush()
            
            # Run analysis
            results = run_kpi_analysis(
                tmp_e.name,
                tmp_w.name,
                products if products else None,
                month if month != 0 else None
            )
            
            # Cleanup temp files
            os.unlink(tmp_e.name)
            os.unlink(tmp_w.name)
            
            if not results['success']:
                st.error(f"❌ Analysis failed: {results['error']}")
            else:
                summary = results['summary']
                yearly = results['yearly']
                
                if yearly.empty:
                    st.warning("⚠️ No data found with selected filters")
                else:
                    # KPI Cards
                    st.markdown("### 📈 Key Performance Indicators")
                    
                    total_energy = yearly["Energy_kWh"].sum()
                    avg_kpi = yearly["kWh_per_m3"].mean()
                    total_volume = yearly["Volume_m3"].sum()
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.markdown(create_kpi_card("Total Energy", total_energy, "kWh"), 
                                   unsafe_allow_html=True)
                    with col2:
                        st.markdown(create_kpi_card("Avg Efficiency", avg_kpi, "kWh/m³"), 
                                   unsafe_allow_html=True)
                    with col3:
                        st.markdown(create_kpi_card("Total Volume", total_volume, "m³"), 
                                   unsafe_allow_html=True)
                    
                    # Charts
                    st.markdown("### 📊 Analysis Charts")
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        fig1 = px.bar(
                            yearly,
                            x="Zone",
                            y="kWh_per_m3",
                            color="Produkt",
                            title="Energy Efficiency by Zone",
                            text_auto=".1f"
                        )
                        fig1.update_layout(height=400)
                        st.plotly_chart(fig1, use_container_width=True)
                    
                    with col2:
                        fig2 = px.pie(
                            yearly,
                            values="Energy_kWh",
                            names="Produkt",
                            title="Energy Distribution"
                        )
                        fig2.update_layout(height=400)
                        st.plotly_chart(fig2, use_container_width=True)
                    
                    # Monthly trend
                    if not summary.empty and "Month" in summary.columns:
                        st.markdown("### 📉 Monthly Trends")
                        fig3 = px.line(
                            summary,
                            x="Month",
                            y="kWh_per_m3",
                            color="Zone",
                            markers=True,
                            title="Efficiency Over Time"
                        )
                        fig3.update_layout(height=400)
                        st.plotly_chart(fig3, use_container_width=True)
                    
                    # Data tables
                    with st.expander("📋 View Data Tables"):
                        tab1, tab2 = st.tabs(["Yearly Summary", "Monthly Detail"])
                        
                        with tab1:
                            st.dataframe(yearly, use_container_width=True)
                        
                        with tab2:
                            st.dataframe(summary, use_container_width=True)
                    
                    # Export
                    st.markdown("### 📥 Export Results")
                    
                    output_file = "KPI_Analysis_Results.xlsx"
                    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
                        yearly.to_excel(writer, sheet_name="Yearly_Summary", index=False)
                        summary.to_excel(writer, sheet_name="Monthly_Detail", index=False)
                    
                    with open(output_file, "rb") as f:
                        st.download_button(
                            "📥 Download Excel Report",
                            f.read(),
                            output_file,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    
                    st.success("✅ Analysis complete!")

else:
    # Instructions
    st.markdown("""
    ## 🚀 How to Use
    
    1. **Upload Files** in the sidebar:
       - Energy consumption file (hourly data)
       - Hordenwagen tracking file
    
    2. **Select Products** to analyze
    
    3. **Optional:** Filter by month
    
    4. **Click "Run Analysis"**
    
    ### 📊 What You'll Get:
    
    - Energy efficiency metrics (kWh/m³)
    - Zone-by-zone analysis
    - Product comparisons
    - Monthly trends
    - Downloadable Excel report
    
    **Note:** For production order optimization, use the separate Optimization app.
    """)
