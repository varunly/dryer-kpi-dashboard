import streamlit as st
import pandas as pd
import tempfile
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
import sys
import pickle
import os
from datetime import datetime, timedelta
import numpy as np
from itertools import permutations

# Import the KPI calculation module
try:
    from dryer_kpi_monthly_final import (
        parse_energy, parse_wagon, explode_intervals, 
        allocate_energy, CONFIG
    )
except ImportError:
    st.error("❌ Unable to import dryer_kpi_monthly_final module")
    st.stop()

# ------------------ Historical Data Manager ------------------
class HistoricalDataManager:
    def __init__(self, storage_path="dryer_historical_data"):
        """Initialize historical data manager"""
        self.storage_path = storage_path
        self.is_first_run = False
        
        if not os.path.exists(storage_path):
            try:
                os.makedirs(storage_path)
                self.is_first_run = True
            except:
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
        except:
            pass
        return []
    
    def get_consolidated_historical_data(self):
        """Get consolidated historical data"""
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
            
            mask = consolidated['Volume_m3'] > 0
            consolidated.loc[mask, 'kWh_per_m3'] = (
                consolidated.loc[mask, 'Energy_kWh'] / 
                consolidated.loc[mask, 'Volume_m3']
            )
            
            return consolidated
        except Exception as e:
            print(f"Error consolidating: {str(e)}")
            return None
    
    def get_status_message(self):
        """Get status message"""
        try:
            history = self.load_kpi_history()
            if not history:
                return "🆕 No historical data yet. Run your first analysis!"
            
            latest = history[-1]['timestamp']
            age = datetime.now() - latest
            
            if len(history) < 5:
                return f"📊 Building history... ({len(history)} analyses)"
            else:
                return f"✅ {len(history)} analyses (latest: {age.days}d ago)"
        except:
            return "📊 Historical data available"

# Initialize historical manager
hdm = HistoricalDataManager()

# ------------------ Production Optimizer ------------------
class ProductionOptimizer:
    def __init__(self, historical_data):
        """Initialize optimizer with KPI data"""
        self.historical_data = historical_data
        
        # Product database with characteristics
        self.PRODUCT_DB = {
            'L28': {'thickness': 28, 'type': 'L'},
            'L30': {'thickness': 30, 'type': 'L'},
            'L32': {'thickness': 32, 'type': 'L'},
            'L34': {'thickness': 34, 'type': 'L'},
            'L36': {'thickness': 36, 'type': 'L'},
            'L38': {'thickness': 38, 'type': 'L'},
            'L40': {'thickness': 40, 'type': 'L'},
            'L44': {'thickness': 44, 'type': 'L'},
            'N40': {'thickness': 40, 'type': 'N'},
            'N44': {'thickness': 44, 'type': 'N'},
            'U36': {'thickness': 36, 'type': 'U'}
        }
    
    def get_product_energy(self, product):
        """Get average energy consumption for a product from historical data"""
        if self.historical_data is None or self.historical_data.empty:
            # Use thickness-based estimate if no data
            thickness = self.PRODUCT_DB.get(product, {}).get('thickness', 36)
            return 80 + (thickness - 28) * 2  # Rough estimate
        
        prod_data = self.historical_data[self.historical_data['Produkt'] == product]
        if not prod_data.empty:
            return prod_data['kWh_per_m3'].mean()
        else:
            thickness = self.PRODUCT_DB.get(product, {}).get('thickness', 36)
            return 80 + (thickness - 28) * 2
    
    def calculate_transition_cost(self, prod1, prod2):
        """Calculate transition cost between two products"""
        p1 = self.PRODUCT_DB.get(prod1, {})
        p2 = self.PRODUCT_DB.get(prod2, {})
        
        if not p1 or not p2:
            return 1000
        
        cost = 0
        
        # 1. Thickness change cost
        thickness_diff = abs(p1['thickness'] - p2['thickness'])
        cost += thickness_diff * 3.0  # 3 kWh per mm difference
        
        # 2. Material type change penalty
        if p1['type'] != p2['type']:
            cost += 50  # Heavy penalty for material change
        
        # 3. Energy consumption difference (from actual data)
        energy1 = self.get_product_energy(prod1)
        energy2 = self.get_product_energy(prod2)
        cost += abs(energy2 - energy1) * 0.5
        
        return cost
    
    def optimize_sequence(self, products):
        """Find optimal production sequence"""
        if len(products) <= 1:
            return products, 0, {}
        
        if len(products) <= 7:
            # Exhaustive search for small sets
            best_order = None
            best_cost = float('inf')
            all_costs = []
            
            for perm in permutations(products):
                cost = sum(
                    self.calculate_transition_cost(perm[i], perm[i+1])
                    for i in range(len(perm)-1)
                )
                all_costs.append(cost)
                if cost < best_cost:
                    best_cost = cost
                    best_order = list(perm)
            
            worst_cost = max(all_costs) if all_costs else best_cost
            avg_cost = np.mean(all_costs) if all_costs else best_cost
            
        else:
            # Intelligent greedy algorithm for larger sets
            best_order = self._intelligent_sequencing(products)
            best_cost = sum(
                self.calculate_transition_cost(best_order[i], best_order[i+1])
                for i in range(len(best_order)-1)
            )
            
            # Estimate worst and average
            random_order = list(products)
            np.random.shuffle(random_order)
            worst_cost = sum(
                self.calculate_transition_cost(random_order[i], random_order[i+1])
                for i in range(len(random_order)-1)
            )
            avg_cost = (best_cost + worst_cost) / 2
        
        metrics = {
            'best_cost': best_cost,
            'worst_cost': worst_cost,
            'avg_cost': avg_cost,
            'savings_vs_worst': (worst_cost - best_cost) / worst_cost if worst_cost > 0 else 0,
            'savings_vs_avg': (avg_cost - best_cost) / avg_cost if avg_cost > 0 else 0
        }
        
        return best_order, best_cost, metrics
    
    def _intelligent_sequencing(self, products):
        """Intelligent sequencing with lookahead"""
        if not products:
            return []
        
        remaining = set(products)
        
        # Start with thinnest product (easier to heat from cold)
        start = min(remaining, 
                   key=lambda p: self.PRODUCT_DB.get(p, {}).get('thickness', 100))
        sequence = [start]
        remaining.remove(start)
        
        # Build sequence with lookahead
        while remaining:
            current = sequence[-1]
            best_next = None
            best_score = float('inf')
            
            for next_prod in remaining:
                immediate = self.calculate_transition_cost(current, next_prod)
                
                # Lookahead
                future = 0
                if len(remaining) > 1:
                    temp_remaining = remaining - {next_prod}
                    if temp_remaining:
                        min_future = min(
                            self.calculate_transition_cost(next_prod, fp)
                            for fp in temp_remaining
                        )
                        future = min_future * 0.3
                
                score = immediate + future
                if score < best_score:
                    best_score = score
                    best_next = next_prod
            
            sequence.append(best_next)
            remaining.remove(best_next)
        
        return sequence
    
    def generate_recommendations(self, sequence, products_demand):
        """Generate optimization recommendations"""
        recommendations = []
        
        # Analyze sequence
        for i in range(len(sequence)-1):
            curr = sequence[i]
            next_prod = sequence[i+1]
            
            curr_info = self.PRODUCT_DB.get(curr, {})
            next_info = self.PRODUCT_DB.get(next_prod, {})
            
            thickness_diff = abs(curr_info.get('thickness', 0) - next_info.get('thickness', 0))
            
            if thickness_diff > 8:
                recommendations.append(
                    f"⚠️ Large thickness jump: {curr} ({curr_info['thickness']}mm) → "
                    f"{next_prod} ({next_info['thickness']}mm). Allow extra setup time."
                )
            
            if curr_info.get('type') != next_info.get('type'):
                recommendations.append(
                    f"🔧 Material change: {curr} → {next_prod}. "
                    f"Schedule cleaning and quality check."
                )
        
        # Energy optimization
        if self.historical_data is not None:
            high_energy = []
            for prod in sequence:
                energy = self.get_product_energy(prod)
                if energy > 100:
                    high_energy.append(prod)
            
            if high_energy:
                recommendations.append(
                    f"💡 High energy products detected: {', '.join(high_energy)}. "
                    f"Consider running during off-peak hours (night)."
                )
        
        # Production efficiency
        total_wagons = sum(products_demand.values())
        if total_wagons > 100:
            recommendations.append(
                f"📊 High volume week ({total_wagons} wagons). "
                f"Consider split production or night shifts."
            )
        
        return recommendations

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
    
    .optimization-result {
        background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
        padding: 20px;
        border-radius: 15px;
        color: white;
        margin: 20px 0;
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

def run_analysis(energy_path, wagon_path, products_filter, month_filter, use_history=True):
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
        
        # Save to history
        results = {
            'summary': summary,
            'yearly': yearly,
            'energy': e,
            'wagons': w,
            'intervals': ivals,
            'allocation': alloc
        }
        
        try:
            if use_history:
                hdm.save_kpi_results(results)
        except:
            pass
        
        progress_bar.empty()
        status_text.empty()
        
        return results
        
    except Exception as e:
        status_text.empty()
        progress_bar.empty()
        raise e

# ------------------ Header ------------------
st.markdown('<div class="main-title">🏭 Lindner – Dryer KPI & Production Optimizer</div>', 
            unsafe_allow_html=True)

st.info("📊 Upload your files to analyze efficiency and get optimized production sequences")

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
    
    # Product selection with select all
    all_products = ["L28", "L30", "L32", "L34", "L36", "L38", "L40", "L44", "N40", "N44", "U36"]
    
    select_all = st.checkbox("Select All Products", value=True, key="select_all_kpi")
    
    if select_all:
        products = st.multiselect(
            "🧱 Product(s):",
            all_products,
            default=all_products,
            help="Select products to analyze",
            key="products_kpi",
            disabled=True
        )
    else:
        products = st.multiselect(
            "🧱 Product(s):",
            all_products,
            default=["L36"],
            help="Select products to analyze",
            key="products_kpi_manual"
        )
    
    st.info(f"📊 Selected: {len(products)} product(s)")
    
    month = st.number_input(
        "📅 Month (0 = all):",
        min_value=0,
        max_value=12,
        value=0,
        help="Filter by specific month"
    )
    
    st.markdown("---")
    st.subheader("📚 Historical Data")
    
    use_historical = st.checkbox(
        "Use Historical Data",
        value=True,
        help="Enhance analysis with historical patterns"
    )
    
    status_message = hdm.get_status_message()
    st.info(status_message)
    
    st.markdown("---")
    run_button = st.button("▶️ Run Analysis", use_container_width=True)

# ------------------ Main Processing ------------------
if run_button:
    if not energy_file or not wagon_file:
        st.error("⚠️ Please upload both files before running analysis.")
    else:
        try:
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
                    month if month != 0 else None,
                    use_historical
                )
                
                # Store in session state
                st.session_state['analysis_results'] = results
                
                summary = results['summary']
                yearly = results['yearly']
                
                if summary.empty:
                    st.warning("⚠️ No data found matching the selected filters.")
                    st.stop()
                
                # --------------- KPI Cards ---------------
                st.markdown('<div class="section-header">📈 Energy Efficiency KPIs</div>', 
                           unsafe_allow_html=True)
                
                total_energy = yearly["Energy_kWh"].sum()
                avg_kpi = yearly["kWh_per_m3"].mean()
                total_volume = yearly["Volume_m3"].sum()
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.markdown(create_kpi_card("Total Energy", total_energy, "kWh"), 
                               unsafe_allow_html=True)
                with col2:
                    st.markdown(create_kpi_card("Avg. Efficiency", avg_kpi, "kWh/m³"), 
                               unsafe_allow_html=True)
                with col3:
                    st.markdown(create_kpi_card("Total Volume", total_volume, "m³"), 
                               unsafe_allow_html=True)
                
                # --------------- Charts ---------------
                st.markdown('<div class="section-header">📊 Performance Analysis</div>', 
                           unsafe_allow_html=True)
                
                col1, col2 = st.columns(2)
                
                with col1:
                    fig1 = px.bar(
                        yearly,
                        x="Zone",
                        y="kWh_per_m3",
                        color="Produkt",
                        text_auto=".1f",
                        title="Energy Efficiency by Zone & Product"
                    )
                    fig1.update_layout(height=400, plot_bgcolor="white")
                    st.plotly_chart(fig1, use_container_width=True)
                
                with col2:
                    fig2 = px.pie(
                        yearly,
                        values="Energy_kWh",
                        names="Produkt",
                        title="Energy Distribution by Product"
                    )
                    fig2.update_layout(height=400)
                    st.plotly_chart(fig2, use_container_width=True)
                
                # --------------- PRODUCTION OPTIMIZATION ---------------
                st.markdown('<div class="section-header">🔄 Production Order Optimization</div>', 
                           unsafe_allow_html=True)
                
                st.info("Based on your KPI data, here's the optimized production sequence:")
                
                # Get products from the analysis
                products_analyzed = yearly['Produkt'].unique().tolist()
                
                # Initialize optimizer with yearly data
                optimizer = ProductionOptimizer(yearly)
                
                # Optimize sequence
                optimal_sequence, total_cost, metrics = optimizer.optimize_sequence(products_analyzed)
                
                # Display results
                col_opt1, col_opt2, col_opt3 = st.columns([2, 1, 1])
                
                with col_opt1:
                    st.markdown("### 🏆 Optimal Production Sequence")
                    sequence_html = " → ".join([f"**{p}**" for p in optimal_sequence])
                    st.markdown(sequence_html)
                    
                    # Show why this sequence is good
                    st.caption("Optimized for minimum energy transitions and setup time")
                
                with col_opt2:
                    st.metric(
                        "Transition Cost",
                        f"{total_cost:.1f} kWh",
                        help="Total energy cost of all transitions"
                    )
                
                with col_opt3:
                    st.metric(
                        "Savings vs Random",
                        f"{metrics['savings_vs_worst']:.1%}",
                        help="Energy saved compared to random sequence"
                    )
                
                # Detailed transition analysis
                with st.expander("📋 Detailed Transition Analysis"):
                    transitions = []
                    for i in range(len(optimal_sequence)-1):
                        from_prod = optimal_sequence[i]
                        to_prod = optimal_sequence[i+1]
                        cost = optimizer.calculate_transition_cost(from_prod, to_prod)
                        
                        from_energy = optimizer.get_product_energy(from_prod)
                        to_energy = optimizer.get_product_energy(to_prod)
                        
                        transitions.append({
                            'From': from_prod,
                            'To': to_prod,
                            'Transition Cost (kWh)': f"{cost:.1f}",
                            'From Energy (kWh/m³)': f"{from_energy:.1f}",
                            'To Energy (kWh/m³)': f"{to_energy:.1f}"
                        })
                    
                    transitions_df = pd.DataFrame(transitions)
                    st.dataframe(transitions_df, use_container_width=True)
                
                # Recommendations
                st.markdown("### 💡 Production Recommendations")
                
                # Dummy demand for recommendations
                products_demand = {p: 10 for p in products_analyzed}  # Assume 10 wagons each
                recommendations = optimizer.generate_recommendations(optimal_sequence, products_demand)
                
                if recommendations:
                    for rec in recommendations:
                        st.info(rec)
                else:
                    st.success("✅ Optimal sequence with smooth transitions!")
                
                # Visualization of sequence
                st.markdown("### 📈 Energy Profile Visualization")
                
                energy_profile = []
                for i, product in enumerate(optimal_sequence):
                    energy = optimizer.get_product_energy(product)
                    energy_profile.append({
                        'Position': i + 1,
                        'Product': product,
                        'Energy (kWh/m³)': energy
                    })
                
                profile_df = pd.DataFrame(energy_profile)
                
                fig3 = px.line(
                    profile_df,
                    x='Position',
                    y='Energy (kWh/m³)',
                    text='Product',
                    markers=True,
                    title="Energy Consumption Profile Through Production Sequence"
                )
                fig3.update_traces(textposition="top center")
                fig3.update_layout(height=400, plot_bgcolor="white")
                st.plotly_chart(fig3, use_container_width=True)
                
                # --------------- Export Section ---------------
                st.markdown('<div class="section-header">📥 Export Results</div>', 
                           unsafe_allow_html=True)
                
                output_path = "Dryer_KPI_Results.xlsx"
                with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
                    results['yearly'].to_excel(writer, sheet_name="KPI_Results", index=False)
                    
                    # Add optimization results
                    opt_df = pd.DataFrame({
                        'Optimal_Sequence': optimal_sequence,
                        'Position': range(1, len(optimal_sequence) + 1)
                    })
                    opt_df.to_excel(writer, sheet_name="Optimal_Sequence", index=False)
                    
                    # Add transition details
                    transitions_df.to_excel(writer, sheet_name="Transitions", index=False)
                
                col_dl1, col_dl2 = st.columns(2)
                
                with col_dl1:
                    with open(output_path, "rb") as f:
                        st.download_button(
                            label="📥 Download Complete Report (Excel)",
                            data=f.read(),
                            file_name="Dryer_KPI_Optimization.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                
                with col_dl2:
                    # Create text report
                    text_report = f"""
LINDNER DRYER - KPI & OPTIMIZATION REPORT
Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}
==========================================

KPI SUMMARY:
- Total Energy: {total_energy:,.2f} kWh
- Average Efficiency: {avg_kpi:.2f} kWh/m³
- Total Volume: {total_volume:.2f} m³

OPTIMAL PRODUCTION SEQUENCE:
{' → '.join(optimal_sequence)}

OPTIMIZATION METRICS:
- Total Transition Cost: {total_cost:.1f} kWh
- Savings vs Worst Case: {metrics['savings_vs_worst']:.1%}
- Savings vs Average: {metrics['savings_vs_avg']:.1%}

RECOMMENDATIONS:
{chr(10).join(['- ' + rec for rec in recommendations])}
                    """
                    
                    st.download_button(
                        label="📄 Download Summary Report (TXT)",
                        data=text_report,
                        file_name="production_optimization_summary.txt",
                        mime="text/plain",
                        use_container_width=True
                    )
                
                st.success("✅ Analysis complete! Use the optimized sequence for your next production run.")
                
        except Exception as e:
            st.error(f"❌ An error occurred: {str(e)}")
            with st.expander("🔍 View Error Details"):
                st.exception(e)

# Show instructions if no analysis has been run
else:
    st.markdown("""
    ## 🚀 How to Use This Tool
    
    1. **Upload Files** in the sidebar:
       - Energy consumption file (hourly data)
       - Hordenwagen tracking file
    
    2. **Select Products** to analyze (or use "Select All")
    
    3. **Click "Run Analysis"** to:
       - Calculate energy efficiency KPIs
       - Get optimized production sequence
       - View recommendations
    
    4. **Download Results** including:
       - KPI analysis
       - Optimal production order
       - Detailed transition analysis
    
    ### 💡 What This Tool Does:
    
    - **Analyzes** your actual energy consumption data
    - **Calculates** efficiency metrics (kWh/m³) for each product and zone
    - **Optimizes** production sequence to minimize:
      - Energy consumption during transitions
      - Setup and changeover time
      - Temperature adjustment costs
    - **Provides** specific recommendations for your operation
    
    ### 📊 The Optimization Considers:
    
    - Actual energy consumption from your data
    - Product thickness differences
    - Material type changes (L, N, U series)
    - Transition costs between products
    
    Start by uploading your files in the sidebar! 👈
    """)
