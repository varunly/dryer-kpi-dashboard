import streamlit as st
import pandas as pd
import tempfile
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
import sys

# Import the KPI calculation module
try:
    from dryer_kpi_monthly_final import main as run_kpi, CONFIG
except ImportError as e:
    st.error("❌ Unable to import dryer_kpi_monthly_final module")
    st.stop()


# ------------------ Page Configuration ------------------
st.set_page_config(
    page_title="Lindner Dryer KPI Dashboard",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded"
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
        text-shadow: 1px 1px 2px rgba(0,0,0,0.1);
    }
    
    .section-header {
        color: #003366;
        font-size: 22px;
        font-weight: 600;
        margin-top: 40px;
        margin-bottom: 20px;
        border-bottom: 2px solid #003366;
        padding-bottom: 6px;
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
    
    .stDownloadButton button {
        background-color: #003366;
        color: white;
        border-radius: 8px;
        padding: 10px 20px;
        font-weight: 600;
    }
    </style>
""", unsafe_allow_html=True)

# ------------------ Header ------------------
st.markdown('<div class="main-title">🏭 Lindner – Dryer KPI Monitoring Dashboard</div>', 
            unsafe_allow_html=True)

st.info("📊 Upload your Energy and Hordenwagen files to analyze dryer efficiency across zones and products.")

# ------------------ Sidebar ------------------
with st.sidebar:
    st.image("https://www.karrieretag.org/wp-content/uploads/2023/10/lindner-logo-1.png", 
             use_column_width=True)
    st.markdown("---")
    
    st.subheader("📁 Data Upload")
    energy_file = st.file_uploader(
        "📊 Energy File (.xlsx)", 
        type=["xlsx"],
        help="Upload the hourly energy consumption Excel file"
    )
    wagon_file = st.file_uploader(
        "🚛 Hordenwagen File (.xlsm, .xlsx)", 
        type=["xlsm", "xlsx"],
        help="Upload the wagon tracking Excel file"
    )
    
    st.markdown("---")
    st.subheader("⚙️ Filters")
    
    products = st.multiselect(
        "🧱 Product(s):",
        ["L30", "L32", "L34", "L36", "L38", "L40", "N40", "N44"],
        default=["L36"],
        help="Select one or more products to analyze"
    )
    
    month = st.number_input(
        "📅 Month (0 = all):",
        min_value=0,
        max_value=12,
        value=0,
        help="Filter by specific month (1-12) or 0 for all months"
    )
    
    st.markdown("---")
    run_button = st.button("▶️ Run Analysis", use_container_width=True)

# ------------------ Helper Functions ------------------
def create_kpi_card(title, value, unit):
    """Create a styled KPI metric card"""
    return f'''
    <div class="metric-card">
        <h3>{title}</h3>
        <h2>{value:,.2f} {unit}</h2>
    </div>
    '''

def run_analysis(energy_path, wagon_path, products_filter, month_filter):
    """Run the KPI analysis with progress tracking"""
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        # Step 1: Parse energy data
        status_text.text("🔄 Parsing energy data...")
        progress_bar.progress(20)
        e_raw = pd.read_excel(energy_path, sheet_name=CONFIG["energy_sheet"])
        e = parse_energy(e_raw)
        
        # Step 2: Parse wagon data
        status_text.text("🔄 Parsing wagon tracking data...")
        progress_bar.progress(40)
        w_raw = pd.read_excel(
            wagon_path, 
            sheet_name=CONFIG["wagon_sheet"], 
            header=CONFIG["wagon_header_row"]
        )
        w = parse_wagon(w_raw)
        
        # Step 3: Apply filters
        status_text.text("🔄 Applying filters...")
        progress_bar.progress(60)
        if products_filter:
            w = w[w["Produkt"].astype(str).isin(products_filter)]
        if month_filter:
            e = e[e["Month"] == month_filter]
            w = w[w["Month"] == month_filter]
        
        # Step 4: Process intervals
        status_text.text("🔄 Processing zone intervals...")
        progress_bar.progress(70)
        ivals = explode_intervals(w)
        
        # Step 5: Allocate energy
        status_text.text("🔄 Allocating energy to products...")
        progress_bar.progress(85)
        alloc = allocate_energy(e, ivals)
        
        # Step 6: Create summaries
        status_text.text("🔄 Generating summaries...")
        progress_bar.progress(95)
        
        summary = alloc.groupby(["Month", "Produkt", "Zone"], as_index=False).agg(
            Energy_kWh=("Energy_share_kWh", "sum"),
            Volume_m3=("m3", "sum")
        )
        summary["kWh_per_m3"] = (
            summary["Energy_kWh"] / summary["Volume_m3"].replace(0, pd.NA)
        )
        
        yearly = summary.groupby(["Produkt", "Zone"], as_index=False).agg(
            Energy_kWh=("Energy_kWh", "sum"),
            Volume_m3=("Volume_m3", "sum")
        )
        yearly["kWh_per_m3"] = (
            yearly["Energy_kWh"] / yearly["Volume_m3"].replace(0, pd.NA)
        )
        
        progress_bar.progress(100)
        status_text.text("✅ Analysis complete!")
        
        return {
            'summary': summary,
            'yearly': yearly,
            'energy': e,
            'wagons': w,
            'intervals': ivals,
            'allocation': alloc
        }
        
    except Exception as e:
        status_text.empty()
        progress_bar.empty()
        raise e

# ------------------ Main Processing ------------------
if run_button:
    if not energy_file or not wagon_file:
        st.error("⚠️ Please upload both files before running analysis.")
    else:
        try:
            # Create temporary files
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_e, \
                 tempfile.NamedTemporaryFile(delete=False, suffix=".xlsm") as tmp_w:
                
                tmp_e.write(energy_file.read())
                tmp_w.write(wagon_file.read())
                tmp_e.flush()
                tmp_w.flush()
                
                # Run analysis
                results = run_analysis(
                    tmp_e.name,
                    tmp_w.name,
                    products if products else None,
                    month if month != 0 else None
                )
                
                summary = results['summary']
                yearly = results['yearly']
                
                # Check if we have data
                if summary.empty:
                    st.warning("⚠️ No data found matching the selected filters.")
                    st.stop()
                
                # --------------- KPI Cards ---------------
                st.markdown('<div class="section-header">📈 Summary KPIs</div>', 
                           unsafe_allow_html=True)
                
                total_energy = yearly["Energy_kWh"].sum()
                avg_kpi = yearly["kWh_per_m3"].mean()
                total_volume = yearly["Volume_m3"].sum()
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.markdown(
                        create_kpi_card("Total Energy", total_energy, "kWh"),
                        unsafe_allow_html=True
                    )
                with col2:
                    st.markdown(
                        create_kpi_card("Avg. Efficiency", avg_kpi, "kWh/m³"),
                        unsafe_allow_html=True
                    )
                with col3:
                    st.markdown(
                        create_kpi_card("Total Volume", total_volume, "m³"),
                        unsafe_allow_html=True
                    )
                
                # --------------- Monthly Trend ---------------
                st.markdown('<div class="section-header">📊 Monthly KPI Trend</div>', 
                           unsafe_allow_html=True)
                
                fig1 = px.line(
                    summary,
                    x="Month",
                    y="kWh_per_m3",
                    color="Zone",
                    markers=True,
                    hover_data=["Produkt", "Energy_kWh", "Volume_m3"],
                    title="Energy Efficiency by Month and Zone"
                )
                fig1.update_layout(
                    height=500,
                    xaxis_title="Month",
                    yaxis_title="kWh/m³",
                    plot_bgcolor="white",
                    hovermode='x unified'
                )
                st.plotly_chart(fig1, use_container_width=True)
                
                # --------------- Zone Comparison ---------------
                st.markdown('<div class="section-header">📉 Zone Comparison</div>', 
                           unsafe_allow_html=True)
                
                col1, col2 = st.columns(2)
                
                with col1:
                    fig2 = px.bar(
                        yearly,
                        x="Zone",
                        y="kWh_per_m3",
                        color="Produkt",
                        text_auto=".2f",
                        title="Yearly KPI by Zone"
                    )
                    fig2.update_layout(height=400, plot_bgcolor="white")
                    st.plotly_chart(fig2, use_container_width=True)
                
                with col2:
                    fig3 = px.pie(
                        yearly,
                        values="Energy_kWh",
                        names="Zone",
                        title="Energy Distribution by Zone"
                    )
                    fig3.update_layout(height=400)
                    st.plotly_chart(fig3, use_container_width=True)
                
                # --------------- Data Tables ---------------
                with st.expander("📋 View Detailed Data Tables"):
                    tab1, tab2 = st.tabs(["Monthly Summary", "Yearly Summary"])
                    
                    with tab1:
                        st.dataframe(
                            summary.style.format({
                                "Energy_kWh": "{:.2f}",
                                "Volume_m3": "{:.2f}",
                                "kWh_per_m3": "{:.2f}"
                            }),
                            use_container_width=True
                        )
                    
                    with tab2:
                        st.dataframe(
                            yearly.style.format({
                                "Energy_kWh": "{:.2f}",
                                "Volume_m3": "{:.2f}",
                                "kWh_per_m3": "{:.2f}"
                            }),
                            use_container_width=True
                        )
                
                # --------------- Download Section ---------------
                st.markdown('<div class="section-header">📥 Export Results</div>', 
                           unsafe_allow_html=True)
                
                # Create Excel file in memory
                output_path = "Dryer_KPI_Results.xlsx"
                with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
                    results['energy'].to_excel(writer, sheet_name="Energy_Data", index=False)
                    results['wagons'].to_excel(writer, sheet_name="Wagon_Data", index=False)
                    results['intervals'].to_excel(writer, sheet_name="Zone_Intervals", index=False)
                    results['allocation'].to_excel(writer, sheet_name="Energy_Allocation", index=False)
                    summary.to_excel(writer, sheet_name="Monthly_Summary", index=False)
                    yearly.to_excel(writer, sheet_name="Yearly_Summary", index=False)
                
                with open(output_path, "rb") as f:
                    st.download_button(
                        label="📥 Download Complete Excel Report",
                        data=f.read(),
                        file_name="Dryer_KPI_Analysis.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
                st.success("✅ Analysis complete! Explore the visualizations above or download the full report.")
                
        except Exception as e:
            st.error(f"❌ An error occurred during analysis: {str(e)}")
            with st.expander("🔍 View Error Details"):
                st.exception(e)


