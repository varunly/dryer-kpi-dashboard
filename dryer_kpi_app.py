import streamlit as st
import pandas as pd
import tempfile
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
import sys
import numpy as np
from itertools import permutations

# Add this complete class to dryer_kpi_app.py (after imports, around line 15)

import pickle
import os
from datetime import datetime
import pandas as pd
import numpy as np

# This is the COMPLETE HistoricalDataManager class with proper try/except blocks
# Replace the ENTIRE class in dryer_kpi_app.py with this:

class HistoricalDataManager:
    def __init__(self, storage_path="dryer_historical_data"):
        """Initialize historical data manager"""
        self.storage_path = storage_path
        self.is_first_run = False
        
        if not os.path.exists(storage_path):
            try:
                os.makedirs(storage_path)
                self.is_first_run = True
            except:  # Added missing except block
                import tempfile
                self.storage_path = tempfile.gettempdir()
        
        self.kpi_file = os.path.join(self.storage_path, "kpi_history.pkl")
        self.optimization_file = os.path.join(self.storage_path, "optimization_history.pkl")
    
    def save_kpi_results(self, results, timestamp=None):
        """Save KPI analysis results with size optimization"""
        try:
            if results is None or 'yearly' not in results:
                return False
                
            if timestamp is None:
                timestamp = datetime.now()
            
            history = self.load_kpi_history()
            
            yearly_data = results.get('yearly')
            
            if yearly_data is not None and not yearly_data.empty:
                # Convert DataFrames to smaller dictionaries
                entry = {
                    'timestamp': timestamp,
                    'products': yearly_data['Produkt'].unique().tolist(),
                    'zones': yearly_data['Zone'].unique().tolist(),
                    'total_energy': float(yearly_data['Energy_kWh'].sum()),
                    'avg_efficiency': float(yearly_data['kWh_per_m3'].mean()),
                    'total_volume': float(yearly_data['Volume_m3'].sum()),
                    'yearly_summary': yearly_data.groupby(['Produkt', 'Zone']).agg({
                        'Energy_kWh': 'sum',
                        'Volume_m3': 'sum',
                        'kWh_per_m3': 'mean'
                    }).reset_index().to_dict('records')
                }
                
                history.append(entry)
                
                # Keep only last 30 entries
                if len(history) > 30:
                    history = history[-30:]
                
                with open(self.kpi_file, 'wb') as f:
                    pickle.dump(history, f)
                
                return True
            return False
            
        except Exception as e:
            st.warning(f"Could not save historical data: {str(e)}")
            return False
    
    def load_kpi_history(self):
        """Load KPI history"""
        try:
            if os.path.exists(self.kpi_file):
                with open(self.kpi_file, 'rb') as f:
                    return pickle.load(f)
        except:  # Added except block
            pass
        return []
    
    def get_consolidated_historical_data(self):
        """Get consolidated historical data for ALL products"""
        try:
            history = self.load_kpi_history()
            
            if not history:
                return None
            
            all_summaries = []
            for entry in history:
                if 'yearly_summary' in entry and entry['yearly_summary']:
                    summary_df = pd.DataFrame(entry['yearly_summary'])
                    summary_df['analysis_date'] = entry.get('timestamp', datetime.now())
                    all_summaries.append(summary_df)
            
            if not all_summaries:
                return None
            
            combined = pd.concat(all_summaries, ignore_index=True)
            
            if combined.empty:
                return None
            
            consolidated = combined.groupby(['Produkt', 'Zone']).agg({
                'Energy_kWh': 'sum',
                'Volume_m3': 'sum',
                'kWh_per_m3': 'mean'
            }).reset_index()
            
            # Recalculate kWh_per_m3
            mask = consolidated['Volume_m3'] > 0
            consolidated.loc[mask, 'kWh_per_m3'] = (
                consolidated.loc[mask, 'Energy_kWh'] / 
                consolidated.loc[mask, 'Volume_m3']
            )
            
            return consolidated
            
        except Exception as e:
            print(f"Error consolidating: {str(e)}")
            return None
    
    def merge_with_current_data(self, current_yearly, weight_historical=0.3):
        """Merge historical data with current analysis"""
        if current_yearly is None or current_yearly.empty:
            return current_yearly, "No current data to merge"
        
        try:
            historical = self.get_consolidated_historical_data()
            
            if historical is None or historical.empty:
                return current_yearly, "No historical data available (using current data only)"
            
            merged = current_yearly.merge(
                historical[['Produkt', 'Zone', 'kWh_per_m3']].add_suffix('_hist'),
                left_on=['Produkt', 'Zone'],
                right_on=['Produkt_hist', 'Zone_hist'],
                how='left'
            )
            
            # Calculate weighted average where historical exists
            has_hist = merged['kWh_per_m3_hist'].notna()
            merged.loc[has_hist, 'kWh_per_m3'] = (
                merged.loc[has_hist, 'kWh_per_m3'] * (1 - weight_historical) +
                merged.loc[has_hist, 'kWh_per_m3_hist'] * weight_historical
            )
            
            # Clean up columns
            merged = merged[['Produkt', 'Zone', 'Energy_kWh', 'Volume_m3', 'kWh_per_m3']]
            
            products_with_history = has_hist.sum()
            status = f"Enhanced with historical data ({products_with_history} records)"
            
            return merged, status
            
        except Exception as e:
            return current_yearly, f"Could not merge: {str(e)}"
    
    def save_optimization_result(self, products, optimal_order, metrics):
        """Save optimization results"""
        try:
            history = self.load_optimization_history()
            
            entry = {
                'timestamp': datetime.now(),
                'products': products,
                'optimal_order': optimal_order,
                'best_cost': metrics.get('best_cost', 0),
                'savings_vs_worst': metrics.get('savings_vs_worst', 0),
                'savings_vs_avg': metrics.get('savings_vs_avg', 0)
            }
            
            history.append(entry)
            
            if len(history) > 30:
                history = history[-30:]
            
            with open(self.optimization_file, 'wb') as f:
                pickle.dump(history, f)
                
        except Exception as e:
            print(f"Could not save optimization: {str(e)}")
    
    def load_optimization_history(self):
        """Load optimization history"""
        try:
            if os.path.exists(self.optimization_file):
                with open(self.optimization_file, 'rb') as f:
                    return pickle.load(f)
        except:  # Added except block
            pass
        return []
    
    def get_status_message(self):
        """Get informative status message about historical data"""
        try:
            history = self.load_kpi_history()
            
            if not history:
                return "🆕 No historical data yet. Run your first analysis to start!"
            
            latest = history[-1]['timestamp']
            age = datetime.now() - latest
            
            if len(history) < 5:
                return f"📊 Building history... ({len(history)} analyses)"
            else:
                return f"✅ {len(history)} analyses (latest: {age.days}d ago)"
        except:  # Added except block
            return "📊 Historical data available"

# Initialize the historical data manager
hdm = HistoricalDataManager()

# Show initial status
if hdm.is_first_run:
    st.info("🎉 Welcome! This appears to be your first run. Historical data tracking will begin with your first analysis.")

# Import the KPI calculation module
try:
    from dryer_kpi_monthly_final import (
        parse_energy, parse_wagon, explode_intervals, 
        allocate_energy, CONFIG
    )
except ImportError:
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
    
    .optimization-card {
        background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
        padding: 15px;
        border-radius: 10px;
        color: white;
        margin-bottom: 10px;
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

# ------------------ Helper Functions for Optimization ------------------
def calculate_product_characteristics(product):
    """Extract product characteristics for optimization"""
    prefix = product[0]  # L, N, or U
    thickness = int(product[1:]) if product[1:].isdigit() else 0
    return prefix, thickness

def calculate_transition_cost(prod1, prod2, historical_data=None):
    """Calculate the cost of transitioning between two products"""
    prefix1, thick1 = calculate_product_characteristics(prod1)
    prefix2, thick2 = calculate_product_characteristics(prod2)
    
    cost = 0
    
    # Thickness change penalty (energy for temperature adjustment)
    thickness_diff = abs(thick2 - thick1)
    cost += thickness_diff * 2.5  # 2.5 kWh per mm difference
    
    # Material type change penalty
    if prefix1 != prefix2:
        cost += 25  # Fixed penalty for material change
    
    # Use historical data if available
    if historical_data is not None and not historical_data.empty:
        prod1_data = historical_data[historical_data['Produkt'] == prod1]
        prod2_data = historical_data[historical_data['Produkt'] == prod2]
        
        if not prod1_data.empty and not prod2_data.empty:
            energy1 = prod1_data['kWh_per_m3'].mean()
            energy2 = prod2_data['kWh_per_m3'].mean()
            if not np.isnan(energy1) and not np.isnan(energy2):
                cost += abs(energy2 - energy1) * 0.8
    
    return cost

def optimize_production_sequence(products, historical_data=None):
    """Find the optimal production sequence to minimize energy consumption"""
    
    if len(products) <= 1:
        return products, 0, {}
    
    # For small lists (≤ 7), check all permutations
    if len(products) <= 7:
        best_order = None
        best_cost = float('inf')
        all_costs = []
        
        for perm in permutations(products):
            cost = sum(calculate_transition_cost(perm[i], perm[i+1], historical_data) 
                      for i in range(len(perm)-1))
            all_costs.append(cost)
            if cost < best_cost:
                best_cost = cost
                best_order = list(perm)
        
        # Calculate savings
        worst_cost = max(all_costs) if all_costs else best_cost
        avg_cost = np.mean(all_costs) if all_costs else best_cost
        
    else:
        # For larger lists, use greedy algorithm with 2-opt improvement
        # Initial greedy solution
        remaining = set(products)
        
        # Start with the product with lowest energy (if data available)
        if historical_data is not None and not historical_data.empty:
            energy_map = {}
            for prod in products:
                prod_data = historical_data[historical_data['Produkt'] == prod]
                energy_map[prod] = prod_data['kWh_per_m3'].mean() if not prod_data.empty else 100
            current = min(remaining, key=lambda x: energy_map.get(x, 100))
        else:
            current = min(remaining, key=lambda x: int(x[1:]) if x[1:].isdigit() else 0)
        
        best_order = [current]
        remaining.remove(current)
        
        # Build sequence greedily
        while remaining:
            next_prod = min(remaining, 
                          key=lambda x: calculate_transition_cost(current, x, historical_data))
            best_order.append(next_prod)
            remaining.remove(next_prod)
            current = next_prod
        
        # Apply 2-opt improvement
        improved = True
        while improved:
            improved = False
            for i in range(1, len(best_order) - 1):
                for j in range(i + 1, len(best_order)):
                    # Try reversing the segment
                    new_order = best_order[:i] + best_order[i:j+1][::-1] + best_order[j+1:]
                    
                    old_cost = sum(calculate_transition_cost(best_order[k], best_order[k+1], historical_data) 
                                 for k in range(len(best_order)-1))
                    new_cost = sum(calculate_transition_cost(new_order[k], new_order[k+1], historical_data) 
                                 for k in range(len(new_order)-1))
                    
                    if new_cost < old_cost:
                        best_order = new_order
                        improved = True
                        break
                if improved:
                    break
        
        best_cost = sum(calculate_transition_cost(best_order[i], best_order[i+1], historical_data) 
                       for i in range(len(best_order)-1))
        
        # Estimate worst case
        worst_order = sorted(products, key=lambda x: int(x[1:]) if x[1:].isdigit() else 0)
        worst_order = worst_order[::2] + worst_order[1::2][::-1]  # Zigzag pattern
        worst_cost = sum(calculate_transition_cost(worst_order[i], worst_order[i+1], historical_data) 
                        for i in range(len(worst_order)-1))
        avg_cost = (best_cost + worst_cost) / 2
    
    metrics = {
        'best_cost': best_cost,
        'worst_cost': worst_cost,
        'avg_cost': avg_cost,
        'savings_vs_worst': (worst_cost - best_cost) / worst_cost if worst_cost > 0 else 0,
        'savings_vs_avg': (avg_cost - best_cost) / avg_cost if avg_cost > 0 else 0
    }
    
    return best_order, best_cost, metrics

def create_sequence_visualization(order, costs, historical_data=None):
    """Create visualization for the production sequence"""
    
    # Get energy values for each product
    energy_values = []
    for product in order:
        if historical_data is not None and not historical_data.empty:
            prod_data = historical_data[historical_data['Produkt'] == product]
            if not prod_data.empty:
                energy_values.append(prod_data['kWh_per_m3'].mean())
            else:
                energy_values.append(100 + int(product[1:]) if product[1:].isdigit() else 100)
        else:
            energy_values.append(100 + int(product[1:]) if product[1:].isdigit() else 100)
    
    # Create figure with subplots
    fig = go.Figure()
    
    # Add main energy consumption trace
    fig.add_trace(go.Scatter(
        x=list(range(len(order))),
        y=energy_values,
        mode='lines+markers+text',
        name='Energy Level',
        text=order,
        textposition="top center",
        line=dict(color='#005691', width=3),
        marker=dict(size=12, color='#003366', 
                   line=dict(color='white', width=2))
    ))
    
    # Add transition cost annotations
    for i in range(len(order)-1):
        transition_cost = calculate_transition_cost(order[i], order[i+1], historical_data)
        
        # Add arrow showing transition
        fig.add_annotation(
            x=i,
            y=energy_values[i],
            ax=i+1,
            ay=energy_values[i+1],
            xref="x",
            yref="y",
            axref="x",
            ayref="y",
            showarrow=True,
            arrowhead=3,
            arrowsize=1,
            arrowwidth=1,
            arrowcolor="red",
            opacity=0.5
        )
        
        # Add cost label
        fig.add_annotation(
            x=i+0.5,
            y=(energy_values[i] + energy_values[i+1])/2,
            text=f"Cost: {transition_cost:.1f}",
            showarrow=False,
            font=dict(size=10, color='red'),
            bgcolor="white",
            opacity=0.8
        )
    
    fig.update_layout(
        title="Production Sequence Energy Profile & Transition Costs",
        xaxis_title="Production Order",
        yaxis_title="Energy Consumption (kWh/m³)",
        height=450,
        showlegend=False,
        plot_bgcolor='white',
        xaxis=dict(
            tickmode='array',
            tickvals=list(range(len(order))),
            ticktext=order,
            showgrid=True,
            gridcolor='lightgray'
        ),
        yaxis=dict(showgrid=True, gridcolor='lightgray')
    )
    
    return fig

# ------------------ Main App Functions ------------------
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

# ------------------ Header ------------------
st.markdown('<div class="main-title">🏭 Lindner – Dryer KPI Monitoring Dashboard</div>', 
            unsafe_allow_html=True)

# Create tabs for different functionalities
tab1, tab2 = st.tabs(["📊 KPI Analysis", "🔄 Production Order Optimization"])

# ------------------ Tab 1: KPI Analysis ------------------
with tab1:
    st.info("📊 Upload your Energy and Hordenwagen files to analyze dryer efficiency across zones and products.")
    
    # Sidebar for KPI Analysis
    with st.sidebar:
        st.image("https://www.karrieretag.org/wp-content/uploads/2023/10/lindner-logo-1.png", 
                 use_column_width=True)
        st.markdown("---")
        
        st.subheader("📁 Data Upload")
        energy_file = st.file_uploader(
            "📊 Energy File (.xlsx)", 
            type=["xlsx"],
            help="Upload the hourly energy consumption Excel file",
            key="energy_kpi"
        )
        wagon_file = st.file_uploader(
            "🚛 Hordenwagen File (.xlsm, .xlsx)", 
            type=["xlsm", "xlsx"],
            help="Upload the wagon tracking Excel file",
            key="wagon_kpi"
        )
        
        st.markdown("---")
        st.subheader("⚙️ Filters")
        
        products = st.multiselect(
            "🧱 Product(s):",
            ["L28", "L30", "L32", "L34", "L36", "L38", "L40", "L44", "N40", "N44", "U36"],
            default=["L36"],
            help="Select one or more products to analyze",
            key="products_kpi"
        )
        
        month = st.number_input(
            "📅 Month (0 = all):",
            min_value=0,
            max_value=12,
            value=0,
            help="Filter by specific month (1-12) or 0 for all months",
            key="month_kpi"
        )
        
        st.markdown("---")
        run_button = st.button("▶️ Run Analysis", use_container_width=True, key="run_kpi")
    
    # Main area for KPI results
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
                    
                    # Store results in session state for optimization tab
                    st.session_state['analysis_results'] = results
                    
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

# ------------------ Tab 2: Production Order Optimization ------------------
with tab2:
    st.markdown('<div class="section-header">🔄 Production Order Optimization</div>', 
                unsafe_allow_html=True)
    
    st.info("Select products to find the most energy-efficient production sequence based on transition costs and historical data.")
    
    col1, col2, col3 = st.columns([2, 1, 1])
    
    with col1:
        available_products = ["L28", "L30", "L32", "L34", "L36", "L38", "L40", "L44", "N40", "N44", "U36"]
        selected_products = st.multiselect(
            "Select products for optimization:",
            available_products,
            default=["L36", "L38", "L30", "L32"],
            help="Choose the products you plan to produce",
            key="opt_products"
        )
    
    with col2:
        wagons_per_product = st.number_input(
            "Wagons per product:",
            min_value=1,
            max_value=100,
            value=20,
            help="Number of wagons for each product type",
            key="wagons_count"
        )
    
    with col3:
        use_historical = st.checkbox(
            "Use historical data",
            value=True,
            help="Use KPI analysis results for better optimization",
            key="use_hist"
        )
    
    # Optimization button
    # FIND THE OPTIMIZATION BUTTON CLICK (if st.button("🔍 Calculate Optimal Sequence"))
# AND REPLACE THE WHOLE BLOCK WITH:

    if st.button("🔍 Calculate Optimal Sequence", use_container_width=True, key="optimize"):
        if len(selected_products) < 2:
            st.error("❌ Please select at least 2 products for optimization.")
        elif len(selected_products) > 4:
            st.error("❌ Please select maximum 4 products for optimization.")
        else:
            with st.spinner("Calculating with historical intelligence..."):
                # Get historical data
                historical_data = hdm.get_consolidated_historical_data()
                
                # Use session state results if available
                current_data = None
                if 'analysis_results' in st.session_state:
                    current_data = st.session_state['analysis_results']['yearly']
                
                # Combine historical and current
                if historical_data is not None:
                    if current_data is not None:
                        st.success("✅ Using combined historical + current data")
                        combined_data, _ = hdm.merge_with_current_data(current_data, weight_historical=0.5)
                    else:
                        st.info("📚 Using historical data")
                        combined_data = historical_data
                elif current_data is not None:
                    st.warning("⚠️ Using current session only")
                    combined_data = current_data
                else:
                    st.error("❌ No data available. Run KPI analysis first.")
                    st.stop()
                
                # Run optimization
                optimal_order, total_cost, metrics = optimize_production_sequence(
                    selected_products, 
                    combined_data  # Now uses historical!
                )
                
                # Save optimization result
                hdm.save_optimization_result(selected_products, optimal_order, metrics)
            
            # ... rest of the display code remains the same ...
                
                # Display results
                st.success("✅ Optimal production sequence calculated!")
                
                # Results in columns
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.markdown("### 🏆 Recommended Order")
                    for i, product in enumerate(optimal_order, 1):
                        st.markdown(f'<div class="optimization-card">{i}. <b>{product}</b></div>', 
                                   unsafe_allow_html=True)
                
                with col2:
                    st.markdown("### 💰 Cost Metrics")
                    st.metric("Total Transition Cost", f"{total_cost:.1f} kWh")
                    st.metric("Avg Cost vs Optimal", f"{metrics['savings_vs_avg']:.1%}")
                    st.metric("Worst Case vs Optimal", f"{metrics['savings_vs_worst']:.1%}")
                
                with col3:
                    st.markdown("### 📊 Production Stats")
                    total_wagons = len(selected_products) * wagons_per_product
                    st.metric("Total Wagons", total_wagons)
                    st.metric("Product Types", len(selected_products))
                    st.metric("Avg Transition Cost", f"{total_cost/(len(optimal_order)-1):.1f} kWh")
                
                # Visualization
                st.markdown("### 📈 Production Sequence Visualization")
                fig_seq = create_sequence_visualization(optimal_order, metrics, historical_data)
                st.plotly_chart(fig_seq, use_container_width=True)
                
                # Recommendations
                st.markdown("### 💡 Production Recommendations")
                
                recommendations = []
                for i in range(len(optimal_order)-1):
                    curr = optimal_order[i]
                    next_prod = optimal_order[i+1]
                    
                    curr_thick = int(curr[1:]) if curr[1:].isdigit() else 0
                    next_thick = int(next_prod[1:]) if next_prod[1:].isdigit() else 0
                    
                    if abs(next_thick - curr_thick) > 6:
                        recommendations.append(
                            f"⚠️ Large thickness change from {curr} ({curr_thick}mm) to {next_prod} ({next_thick}mm). "
                            f"Consider intermediate thickness if available."
                        )
                    
                    if curr[0] != next_prod[0]:
                        recommendations.append(
                            f"🔄 Material type change from {curr} to {next_prod}. "
                            f"Schedule cleaning/adjustment time between batches."
                        )
                
                if recommendations:
                    for rec in recommendations:
                        st.info(rec)
                else:
                    st.success("✅ Sequence is well-optimized with smooth transitions!")
                
                # Export optimization results
                st.markdown("### 📥 Export Optimization Plan")
                
                # Create optimization report
                opt_df = pd.DataFrame({
                    'Order': range(1, len(optimal_order) + 1),
                    'Product': optimal_order,
                    'Wagons': wagons_per_product,
                    'Transition_Cost': [0] + [calculate_transition_cost(optimal_order[i], optimal_order[i+1], historical_data) 
                                              for i in range(len(optimal_order)-1)]
                })
                
                csv = opt_df.to_csv(index=False)
                st.download_button(
                    label="📥 Download Production Plan (CSV)",
                    data=csv,
                    file_name="optimal_production_sequence.csv",
                    mime="text/csv",
                    use_container_width=True
                )
    
    







