"""
Lindner Dryer - KPI Analysis Dashboard
Analyzes energy efficiency and creates reports
"""
# Add at the very top after imports
import streamlit as st
import pandas as pd
import tempfile
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import os
import sys
import numpy as np
from itertools import permutations

# Fix import path for Streamlit Cloud
current_file = os.path.abspath(__file__)
current_dir = os.path.dirname(current_file)
parent_dir = os.path.dirname(current_dir)

# Add paths
sys.path.insert(0, parent_dir)
sys.path.insert(0, os.path.join(parent_dir, 'core'))

# Show debug info
print(f"Current file: {current_file}")
print(f"Current dir: {current_dir}")
print(f"Parent dir: {parent_dir}")
print(f"Python path: {sys.path}")

# Try importing
try:
    # Method 1: Direct import from core
    import core.dryer_kpi_monthly_final as kpi_module
    parse_energy = kpi_module.parse_energy
    parse_wagon = kpi_module.parse_wagon
    explode_intervals = kpi_module.explode_intervals
    allocate_energy = kpi_module.allocate_energy
    CONFIG = kpi_module.CONFIG
    print("✅ Import successful: Method 1 (core.module)")
    
except ImportError as e1:
    print(f"Method 1 failed: {e1}")
    try:
        # Method 2: Add to path and import
        from dryer_kpi_monthly_final import (
            parse_energy, parse_wagon, explode_intervals, 
            allocate_energy, CONFIG
        )
        print("✅ Import successful: Method 2 (direct)")
        
    except ImportError as e2:
        print(f"Method 2 failed: {e2}")
        
        # Show detailed error
        st.error(f"""
        ❌ Cannot import required modules
        
        **Debug Info:**
        - Current file: `{current_file}`
        - Parent dir: `{parent_dir}`
        - Error 1: {e1}
        - Error 2: {e2}
        
        **Files in parent directory:**
        """)
        
        try:
            files = os.listdir(parent_dir)
            st.write(files)
            
            if 'core' in files:
                core_files = os.listdir(os.path.join(parent_dir, 'core'))
                st.write("Files in core/:", core_files)
        except:
            pass
        
        st.stop()

# Add parent directory to path for imports
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

try:
    from core.dryer_kpi_monthly_final import (
        parse_energy, parse_wagon, explode_intervals, 
        allocate_energy, CONFIG
    )
except ImportError:
    try:
        # Fallback for local development
        from dryer_kpi_monthly_final import (
            parse_energy, parse_wagon, explode_intervals, 
            allocate_energy, CONFIG
        )
    except ImportError:
        st.error("❌ Cannot import required modules. Please check file structure.")
        st.stop()

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
    """Run KPI analysis with error handling"""
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        # Step 1: Parse energy data
        status_text.text("🔄 Parsing energy data...")
        progress_bar.progress(20)
        
        try:
            e_raw = pd.read_excel(energy_path, sheet_name=CONFIG["energy_sheet"])
            e = parse_energy(e_raw)
            st.sidebar.success(f"✅ Loaded {len(e)} energy records")
        except Exception as e_error:
            return {'success': False, 'error': f"Energy file error: {str(e_error)}"}
        
        # Step 2: Parse wagon data
        status_text.text("🔄 Parsing wagon tracking data...")
        progress_bar.progress(40)
        
        try:
            w_raw = pd.read_excel(
                wagon_path, 
                sheet_name=CONFIG["wagon_sheet"], 
                header=CONFIG["wagon_header_row"]
            )
            w = parse_wagon(w_raw)
            st.sidebar.success(f"✅ Loaded {len(w)} wagon records")
        except Exception as w_error:
            return {'success': False, 'error': f"Wagon file error: {str(w_error)}"}
        
        # Step 3: Apply filters
        status_text.text("🔄 Applying filters...")
        progress_bar.progress(60)
        
        if products_filter:
            w = w[w["Produkt"].astype(str).isin(products_filter)]
            st.sidebar.info(f"Filtered to {len(w)} wagon records")
        
        if month_filter:
            e = e[e["Month"] == month_filter]
            w = w[w["Month"] == month_filter]
            st.sidebar.info(f"Filtered to month {month_filter}")
        
        if len(w) == 0:
            return {'success': False, 'error': 'No wagon data after filtering'}
        
        # Step 4: Process intervals
        status_text.text("🔄 Processing intervals...")
        progress_bar.progress(70)
        
        try:
            ivals = explode_intervals(w)
            if len(ivals) == 0:
                return {'success': False, 'error': 'No intervals created from wagon data'}
        except Exception as i_error:
            return {'success': False, 'error': f"Interval processing error: {str(i_error)}"}
        
        # Step 5: Allocate energy
        status_text.text("🔄 Allocating energy...")
        progress_bar.progress(85)
        
        try:
            alloc = allocate_energy(e, ivals)
            if len(alloc) == 0:
                return {'success': False, 'error': 'No energy allocation results'}
        except Exception as a_error:
            return {'success': False, 'error': f"Energy allocation error: {str(a_error)}"}
        
        # Step 6: Create summaries
        status_text.text("🔄 Creating summaries...")
        progress_bar.progress(95)
        
        try:
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
        except Exception as s_error:
            return {'success': False, 'error': f"Summary creation error: {str(s_error)}"}
        
        progress_bar.progress(100)
        status_text.text("✅ Analysis complete!")
        
        # Clear progress indicators
        import time
        time.sleep(0.5)
        progress_bar.empty()
        status_text.empty()
        
        return {
            'summary': summary,
            'yearly': yearly,
            'success': True
        }
        
    except Exception as e:
        progress_bar.empty()
        status_text.empty()
        return {'success': False, 'error': f"Unexpected error: {str(e)}"}

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
    energy_file = st.file_uploader("Energy File (.xlsx)", type=["xlsx"], key="energy_upload")
    wagon_file = st.file_uploader("Hordenwagen File (.xlsm, .xlsx)", type=["xlsm", "xlsx"], key="wagon_upload")
    
    st.markdown("---")
    st.subheader("⚙️ Filters")
    
    all_products = ["L28", "L30", "L32", "L34", "L36", "L38", "L40", "L44", "N40", "N44", "U36"]
    
    select_all = st.checkbox("Select All Products", value=True, key="select_all")
    
    if select_all:
        products = all_products
        st.multiselect("Products:", all_products, default=all_products, disabled=True, key="prod_disabled")
    else:
        products = st.multiselect("Products:", all_products, default=["L36"], key="prod_manual")
    
    st.info(f"Selected: {len(products)} products")
    
    month = st.number_input("Month (0 = all):", 0, 12, 0, key="month_filter")
    
    st.markdown("---")
    run_button = st.button("▶️ Run Analysis", use_container_width=True, type="primary")

# ------------------ Main Analysis ------------------
if run_button:
    if not energy_file or not wagon_file:
        st.error("⚠️ Please upload both files")
    else:
        # Clear any previous results
        if 'results' in st.session_state:
            del st.session_state['results']
        
        # Create temporary files
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_e, \
                 tempfile.NamedTemporaryFile(delete=False, suffix=".xlsm") as tmp_w:
                
                # Write uploaded files to temp
                tmp_e.write(energy_file.read())
                tmp_w.write(wagon_file.read())
                tmp_e.flush()
                tmp_w.flush()
                
                # Run analysis
                with st.spinner("Running analysis..."):
                    results = run_kpi_analysis(
                        tmp_e.name,
                        tmp_w.name,
                        products if products else None,
                        month if month != 0 else None
                    )
                
                # Cleanup temp files
                try:
                    os.unlink(tmp_e.name)
                    os.unlink(tmp_w.name)
                except:
                    pass
                
                # Check results
                if not results['success']:
                    st.error(f"❌ Analysis failed: {results['error']}")
                    
                    with st.expander("🔍 Troubleshooting Tips"):
                        st.markdown("""
                        **Common Issues:**
                        
                        1. **Energy file error**: Check that your energy file has the correct format
                        2. **Wagon file error**: Ensure the Hordenwagen file has data starting at row 7
                        3. **No data after filtering**: Try selecting all products and all months
                        4. **No intervals created**: Check that wagon data has valid timestamps
                        
                        **Quick Fixes:**
                        - Select "All Products"
                        - Set Month to "0" (all months)
                        - Verify your files are not corrupted
                        - Check that files contain data for 2025
                        """)
                else:
                    summary = results['summary']
                    yearly = results['yearly']
                    
                    if yearly.empty:
                        st.warning("⚠️ No data found with selected filters")
                        st.info("Try selecting all products and all months")
                    else:
                        # Store in session state
                        st.session_state['results'] = results
                        
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
                            fig1.update_layout(height=400, plot_bgcolor='white')
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
                        
                        # Monthly trend (if available)
                        if not summary.empty and "Month" in summary.columns and summary["Month"].nunique() > 1:
                            st.markdown("### 📉 Monthly Trends")
                            fig3 = px.line(
                                summary,
                                x="Month",
                                y="kWh_per_m3",
                                color="Zone",
                                markers=True,
                                title="Efficiency Over Time"
                            )
                            fig3.update_layout(height=400, plot_bgcolor='white')
                            st.plotly_chart(fig3, use_container_width=True)
                        
                        # Data tables
                        with st.expander("📋 View Data Tables"):
                            tab1, tab2 = st.tabs(["Yearly Summary", "Monthly Detail"])
                            
                            with tab1:
                                st.dataframe(
                                    yearly.style.format({
                                        "Energy_kWh": "{:,.2f}",
                                        "Volume_m3": "{:,.2f}",
                                        "kWh_per_m3": "{:,.2f}"
                                    }),
                                    use_container_width=True
                                )
                            
                            with tab2:
                                st.dataframe(
                                    summary.style.format({
                                        "Energy_kWh": "{:,.2f}",
                                        "Volume_m3": "{:,.2f}",
                                        "kWh_per_m3": "{:,.2f}"
                                    }),
                                    use_container_width=True
                                )
                        
                        # Export
                        st.markdown("### 📥 Export Results")
                        
                        try:
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
                            
                            # Clean up output file
                            if os.path.exists(output_file):
                                os.unlink(output_file)
                                
                        except Exception as export_error:
                            st.warning(f"Could not create download file: {str(export_error)}")
                        
                        st.success("✅ Analysis complete!")
        
        except Exception as e:
            st.error(f"❌ Unexpected error: {str(e)}")
            
            with st.expander("🔍 Error Details"):
                st.exception(e)

else:
    # Instructions
    st.markdown("""
    ## 🚀 How to Use
    
    1. **Upload Files** in the sidebar:
       - Energy consumption file (hourly data)
       - Hordenwagen tracking file
    
    2. **Select Products** to analyze (or use "Select All")
    
    3. **Optional:** Filter by month (0 = all months)
    
    4. **Click "Run Analysis"**
    
    ### 📊 What You'll Get:
    
    - ✅ Energy efficiency metrics (kWh/m³)
    - ✅ Zone-by-zone analysis
    - ✅ Product comparisons
    - ✅ Monthly trends
    - ✅ Downloadable Excel report
    
    ### 📁 File Requirements:
    
    **Energy File (.xlsx):**
    - Hourly energy consumption data
    - Columns: Zeitstempel, Gasmenge Zone 2-5, etc.
    
    **Hordenwagen File (.xlsm or .xlsx):**
    - Wagon tracking data
    - Header row should start at row 7
    - Columns: WG-Nr, Produkt, Press-Zeit, etc.
    
    ---
    
    **Need production order optimization?**  
    Use the separate [Optimization App](https://dryer-production-optimizer.streamlit.app)
    """)

# Add debug info in sidebar
with st.sidebar:
    st.markdown("---")
    if st.checkbox("Show Debug Info", value=False):
        st.write("**Session State:**")
        st.write(f"Has results: {'results' in st.session_state}")
        
        st.write("**File Upload Status:**")
        st.write(f"Energy file: {'✅' if energy_file else '❌'}")
        st.write(f"Wagon file: {'✅' if wagon_file else '❌'}")


