"""
Lindner Dryer - Production Order Optimizer
WITH DYNAMIC ENERGY CALCULATION FROM UPLOADED DATA
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import json
import os
import numpy as np
from itertools import permutations

# ------------------ Page Configuration ------------------
st.set_page_config(
    page_title="Lindner Dryer - Production Optimizer",
    page_icon="🔄",
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

.sequence-box {
    background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
    padding: 30px;
    border-radius: 15px;
    color: white;
    text-align: center;
    font-size: 24px;
    font-weight: 600;
    margin: 20px 0;
}

.metric-card {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    padding: 20px;
    border-radius: 15px;
    text-align: center;
    color: white;
}
</style>
""", unsafe_allow_html=True)

# ------------------ DATA PROCESSOR CLASS ------------------
class ProductionDataProcessor:
    """Process uploaded production data to calculate energy metrics"""
    
    def __init__(self):
        self.df = None
        self.product_profiles = {}
        self.transition_matrix = {}
    
    def load_data(self, uploaded_file):
        """Load and validate uploaded data"""
        try:
            if uploaded_file.name.endswith('.csv'):
                self.df = pd.read_csv(uploaded_file)
            elif uploaded_file.name.endswith(('.xls', '.xlsx')):
                self.df = pd.read_excel(uploaded_file)
            else:
                return False, "Unsupported file format. Please upload CSV or Excel."
            
            # Validate required columns
            required_cols = ['product', 'energy_kwh', 'volume_m3']
            missing_cols = [col for col in required_cols if col not in self.df.columns]
            
            if missing_cols:
                return False, f"Missing required columns: {', '.join(missing_cols)}"
            
            return True, "Data loaded successfully"
        
        except Exception as e:
            return False, f"Error loading file: {str(e)}"
    
    def calculate_energy_metrics(self):
        """Calculate energy consumption per product from uploaded data"""
        
        if self.df is None:
            return False, "No data loaded"
        
        # Group by product and calculate metrics
        product_stats = self.df.groupby('product').agg({
            'energy_kwh': ['sum', 'mean', 'std', 'count'],
            'volume_m3': ['sum', 'mean'],
        }).reset_index()
        
        product_stats.columns = ['_'.join(col).strip('_') for col in product_stats.columns]
        
        # Calculate energy efficiency (kWh per m³)
        for _, row in product_stats.iterrows():
            product = row['product']
            
            total_energy = row['energy_kwh_sum']
            total_volume = row['volume_m3_sum']
            avg_energy = row['energy_kwh_mean']
            batches = row['energy_kwh_count']
            
            # Energy per cubic meter
            kwh_per_m3 = total_energy / total_volume if total_volume > 0 else 0
            
            # Estimate wagon energy (assuming ~9 m³ per wagon - adjust as needed)
            wagon_volume = 9.0  # m³
            kwh_per_wagon = kwh_per_m3 * wagon_volume
            
            # Get additional info if available
            thickness = self.df[self.df['product'] == product]['thickness_mm'].iloc[0] \
                       if 'thickness_mm' in self.df.columns else self._estimate_thickness(product)
            
            product_type = self.df[self.df['product'] == product]['type'].iloc[0] \
                          if 'type' in self.df.columns else self._classify_product_type(product)
            
            self.product_profiles[product] = {
                'type': product_type,
                'thickness_mm': thickness,
                'avg_kwh_per_m3': round(kwh_per_m3, 2),
                'kwh_per_wagon': round(kwh_per_wagon, 1),
                'total_batches': int(batches),
                'total_energy': round(total_energy, 2),
                'total_volume': round(total_volume, 2),
                'std_dev': round(row['energy_kwh_std'], 2)
            }
        
        return True, f"Calculated metrics for {len(self.product_profiles)} products"
    
    def calculate_transition_matrix(self):
        """Calculate transition costs between products"""
        
        products = list(self.product_profiles.keys())
        
        # Initialize matrix
        for p1 in products:
            self.transition_matrix[p1] = {}
            for p2 in products:
                if p1 == p2:
                    self.transition_matrix[p1][p2] = 0
                else:
                    # Calculate transition cost based on:
                    # 1. Thickness difference
                    # 2. Type change
                    # 3. Energy level difference
                    
                    prof1 = self.product_profiles[p1]
                    prof2 = self.product_profiles[p2]
                    
                    thickness_cost = abs(prof2['thickness_mm'] - prof1['thickness_mm']) * 10
                    type_cost = 150 if prof1['type'] != prof2['type'] else 0
                    energy_cost = abs(prof2['avg_kwh_per_m3'] - prof1['avg_kwh_per_m3']) * 0.5
                    
                    total_cost = thickness_cost + type_cost + energy_cost
                    self.transition_matrix[p1][p2] = round(total_cost, 2)
        
        return True, "Transition matrix calculated"
    
    def _estimate_thickness(self, product_name):
        """Estimate thickness from product name (e.g., WS08 -> 8mm)"""
        import re
        match = re.search(r'(\d+)', product_name)
        return int(match.group(1)) if match else 10
    
    def _classify_product_type(self, product_name):
        """Classify product type from name"""
        name_upper = product_name.upper()
        if 'WS' in name_upper or 'WOOD' in name_upper:
            return 'Wood Shavings'
        elif 'HS' in name_upper or 'HEMP' in name_upper:
            return 'Hemp Shavings'
        elif 'SS' in name_upper or 'STRAW' in name_upper:
            return 'Straw'
        else:
            return 'Unknown'
    
    def get_database(self):
        """Return database dict for optimizer"""
        return {
            'product_profiles': self.product_profiles,
            'transition_matrix': self.transition_matrix,
            'metadata': {
                'created': datetime.now().isoformat(),
                'source': 'uploaded_data',
                'products_count': len(self.product_profiles)
            }
        }

# ------------------ OPTIMIZER CLASS ------------------
class ProductionOptimizer:
    def __init__(self, database):
        """Initialize with database dict"""
        self.db = database
        self.profiles = database['product_profiles']
        self.transitions = database['transition_matrix']
        self.rules = database.get('optimization_rules', {})
    
    def optimize(self, products, wagons_per_product=None):
        """Find optimal production sequence"""
        if not products or len(products) < 1:
            return {"error": "No products specified"}
        
        if len(products) == 1:
            return {
                "optimal_sequence": products,
                "total_transition_cost": 0,
                "worst_case_cost": 0,
                "savings_percent": 0,
                "transitions": [],
                "recommendations": ["Single product - no optimization needed"],
                "estimated_total_energy": None
            }
        
        # Find optimal sequence
        if len(products) <= 8:
            best_seq, best_cost = self._exhaustive_search(products)
        else:
            best_seq, best_cost = self._greedy_search(products)
        
        # Calculate worst case
        worst_seq = list(reversed(best_seq))
        worst_cost = self._calculate_cost(worst_seq)
        
        savings = ((worst_cost - best_cost) / worst_cost * 100) if worst_cost > 0 else 0
        
        # Build transition details
        transitions = []
        for i in range(len(best_seq) - 1):
            from_prod = best_seq[i]
            to_prod = best_seq[i+1]
            
            transitions.append({
                "from": from_prod,
                "to": to_prod,
                "cost_kwh": self.transitions[from_prod][to_prod],
                "thickness_change": self.profiles[to_prod]['thickness_mm'] - self.profiles[from_prod]['thickness_mm'],
                "type_change": self.profiles[from_prod]['type'] != self.profiles[to_prod]['type'],
                "energy_change": self.profiles[to_prod]['avg_kwh_per_m3'] - self.profiles[from_prod]['avg_kwh_per_m3']
            })
        
        # Generate recommendations
        recommendations = self._generate_recommendations(transitions, wagons_per_product)
        
        # Estimate energy
        estimated_energy = None
        if wagons_per_product:
            production_energy = sum(
                self.profiles[p]['kwh_per_wagon'] * wagons_per_product.get(p, 0)
                for p in best_seq
            )
            estimated_energy = {
                "production_kwh": round(production_energy, 2),
                "transition_kwh": round(best_cost, 2),
                "total_kwh": round(production_energy + best_cost, 2)
            }
        
        return {
            "optimal_sequence": best_seq,
            "total_transition_cost": round(best_cost, 2),
            "worst_case_cost": round(worst_cost, 2),
            "savings_percent": round(savings, 1),
            "transitions": transitions,
            "recommendations": recommendations,
            "estimated_total_energy": estimated_energy
        }
    
    def _exhaustive_search(self, products):
        """Try all permutations for small sets"""
        best_seq = None
        best_cost = float('inf')
        
        for perm in permutations(products):
            cost = self._calculate_cost(perm)
            if cost < best_cost:
                best_cost = cost
                best_seq = list(perm)
        
        return best_seq, best_cost
    
    def _greedy_search(self, products):
        """Greedy nearest neighbor for larger sets"""
        remaining = set(products)
        
        # Start with thinnest product
        current = min(remaining, key=lambda p: self.profiles[p]['thickness_mm'])
        sequence = [current]
        remaining.remove(current)
        
        # Build sequence greedily
        while remaining:
            next_prod = min(remaining, key=lambda p: self.transitions[current][p])
            sequence.append(next_prod)
            remaining.remove(next_prod)
            current = next_prod
        
        return sequence, self._calculate_cost(sequence)
    
    def _calculate_cost(self, sequence):
        """Calculate total transition cost"""
        if len(sequence) < 2:
            return 0
        return sum(
            self.transitions[sequence[i]][sequence[i+1]]
            for i in range(len(sequence)-1)
        )
    
    def _generate_recommendations(self, transitions, wagons_per_product):
        """Generate recommendations"""
        recs = []
        
        for trans in transitions:
            if trans['cost_kwh'] > 100:
                recs.append(
                    f"⚠️ High transition cost: {trans['from']} → {trans['to']} "
                    f"({trans['cost_kwh']:.1f} kWh). Allow extra setup time."
                )
            
            if trans['type_change']:
                recs.append(
                    f"🔧 Material change: {trans['from']} → {trans['to']}. "
                    f"Schedule cleaning and quality check."
                )
            
            if abs(trans['thickness_change']) > 8:
                recs.append(
                    f"📏 Large thickness change: {trans['from']} → {trans['to']} "
                    f"({trans['thickness_change']:+d}mm). Monitor dryer settings."
                )
        
        if wagons_per_product:
            total = sum(wagons_per_product.values())
            if total > 100:
                recs.append(
                    f"📊 High volume week ({total} wagons). "
                    f"Consider night shifts or split batches."
                )
        
        return recs if recs else ["✅ Optimal sequence with smooth transitions!"]
    
    def get_product_info(self, product):
        """Get product profile"""
        return self.profiles.get(product, {})

# ------------------ Header ------------------
st.markdown('<div class="main-title">🔄 Lindner – Dryer Production Optimizer</div>',
            unsafe_allow_html=True)

# ------------------ Initialize Session State ------------------
if 'processor' not in st.session_state:
    st.session_state.processor = ProductionDataProcessor()
if 'data_loaded' not in st.session_state:
    st.session_state.data_loaded = False
if 'database' not in st.session_state:
    st.session_state.database = None

# ------------------ FILE UPLOAD SECTION ------------------
st.markdown("## 📤 Step 1: Upload Production Data")

st.info("""
**Required columns in your data file:**
- `product` - Product code (e.g., WS08, WS10, HS12)
- `energy_kwh` - Energy consumed per batch/wagon
- `volume_m3` - Volume produced per batch/wagon

**Optional columns:**
- `thickness_mm` - Product thickness
- `type` - Product type (Wood Shavings, Hemp, etc.)
- `date` - Production date
""")

uploaded_file = st.file_uploader(
    "Upload your production data (CSV or Excel)",
    type=['csv', 'xlsx', 'xls'],
    help="File should contain columns: product, energy_kwh, volume_m3"
)

if uploaded_file is not None:
    with st.spinner("Processing data..."):
        # Load data
        success, message = st.session_state.processor.load_data(uploaded_file)
        
        if success:
            st.success(f"✅ {message}")
            
            # Show preview
            st.markdown("### 📊 Data Preview")
            st.dataframe(st.session_state.processor.df.head(10), use_container_width=True)
            
            # Calculate metrics
            calc_success, calc_message = st.session_state.processor.calculate_energy_metrics()
            if calc_success:
                st.success(f"✅ {calc_message}")
                
                # Calculate transitions
                trans_success, trans_message = st.session_state.processor.calculate_transition_matrix()
                if trans_success:
                    st.success(f"✅ {trans_message}")
                    
                    # Store database
                    st.session_state.database = st.session_state.processor.get_database()
                    st.session_state.data_loaded = True
                    
                    # Show calculated energy metrics
                    st.markdown("### ⚡ Calculated Energy Metrics")
                    
                    metrics_data = []
                    for product, profile in st.session_state.database['product_profiles'].items():
                        metrics_data.append({
                            'Product': product,
                            'Type': profile['type'],
                            'Thickness (mm)': profile['thickness_mm'],
                            'kWh/m³': profile['avg_kwh_per_m3'],
                            'kWh/Wagon': profile['kwh_per_wagon'],
                            'Batches': profile['total_batches'],
                            'Total Energy': f"{profile['total_energy']:.0f} kWh",
                            'Std Dev': profile['std_dev']
                        })
                    
                    metrics_df = pd.DataFrame(metrics_data)
                    st.dataframe(metrics_df, use_container_width=True)
                else:
                    st.error(trans_message)
            else:
                st.error(calc_message)
        else:
            st.error(f"❌ {message}")

# ------------------ OPTIMIZATION SECTION ------------------
if st.session_state.data_loaded and st.session_state.database:
    
    st.markdown("---")
    st.markdown("## 🎯 Step 2: Plan Production & Optimize")
    
    # Initialize optimizer
    optimizer = ProductionOptimizer(st.session_state.database)
    
    # Sidebar for production planning
    with st.sidebar:
        st.image("https://www.karrieretag.org/wp-content/uploads/2023/10/lindner-logo-1.png",
                 use_container_width=True)
        st.markdown("---")
        
        st.subheader("📦 Weekly Production Plan")
        st.write("Enter wagons needed per product:")
        
        all_products = list(st.session_state.database['product_profiles'].keys())
        
        weekly_demand = {}
        
        for product in sorted(all_products):
            wagons = st.number_input(
                f"{product}:",
                min_value=0,
                max_value=100,
                value=0,
                key=f"wagon_{product}"
            )
            if wagons > 0:
                weekly_demand[product] = wagons
        
        st.markdown("---")
        
        if weekly_demand:
            total_wagons = sum(weekly_demand.values())
            st.metric("Total Wagons", total_wagons)
            st.metric("Products", len(weekly_demand))
        
        st.markdown("---")
        optimize_button = st.button("🚀 Optimize Production Order", use_container_width=True, type="primary")
    
    # Main optimization results
    if optimize_button:
        if not weekly_demand:
            st.warning("⚠️ Please enter production quantities for at least one product")
        else:
            with st.spinner("🔄 Calculating optimal sequence..."):
                
                products_to_optimize = list(weekly_demand.keys())
                result = optimizer.optimize(products_to_optimize, weekly_demand)
                
                if 'error' in result:
                    st.error(f"❌ {result['error']}")
                else:
                    # Display optimal sequence
                    st.markdown("### 🏆 Optimal Production Sequence")
                    
                    sequence_html = f'''
                    <div class="sequence-box">
                        {' → '.join(result['optimal_sequence'])}
                    </div>
                    '''
                    st.markdown(sequence_html, unsafe_allow_html=True)
                    
                    # Metrics
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        st.metric(
                            "Transition Cost",
                            f"{result['total_transition_cost']:.1f} kWh",
                            help="Energy cost of all transitions"
                        )
                    
                    with col2:
                        st.metric(
                            "Savings",
                            f"{result['savings_percent']:.1f}%",
                            delta="vs worst case",
                            help="Energy saved vs random sequence"
                        )
                    
                    with col3:
                        if result['estimated_total_energy']:
                            st.metric(
                                "Total Energy",
                                f"{result['estimated_total_energy']['total_kwh']:,.0f} kWh",
                                help="Production + transition energy"
                            )
                    
                    # Transition details
                    st.markdown("### 📋 Transition Analysis")
                    
                    transitions_data = []
                    for trans in result['transitions']:
                        transitions_data.append({
                            'From': trans['from'],
                            'To': trans['to'],
                            'Cost (kWh)': f"{trans['cost_kwh']:.1f}",
                            'Thickness Δ (mm)': f"{trans['thickness_change']:+d}",
                            'Type Change': '✓' if trans['type_change'] else '',
                            'Energy Δ (kWh/m³)': f"{trans['energy_change']:+.1f}"
                        })
                    
                    transitions_df = pd.DataFrame(transitions_data)
                    st.dataframe(transitions_df, use_container_width=True)
                    
                    # Visualization
                    st.markdown("### 📈 Energy Profile")
                    
                    energy_profile = []
                    for i, product in enumerate(result['optimal_sequence']):
                        profile = optimizer.get_product_info(product)
                        energy_profile.append({
                            'Position': i + 1,
                            'Product': product,
                            'Energy (kWh/m³)': profile['avg_kwh_per_m3'],
                            'Wagons': weekly_demand.get(product, 0)
                        })
                    
                    profile_df = pd.DataFrame(energy_profile)
                    
                    fig = px.line(
                        profile_df,
                        x='Position',
                        y='Energy (kWh/m³)',
                        text='Product',
                        markers=True,
                        title="Energy Consumption Through Production Sequence"
                    )
                    fig.update_traces(textposition="top center", line=dict(width=3))
                    fig.update_layout(height=400, plot_bgcolor='white')
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # Recommendations
                    st.markdown("### 💡 Production Recommendations")
                    
                    for rec in result['recommendations']:
                        st.info(rec)
                    
                    st.success("✅ Optimization complete!")

else:
    st.markdown("---")
    st.info("👆 Please upload your production data file to begin")

st.markdown("---")
st.caption("🏭 Lindner Dryer - Production Optimizer v3.0 (Dynamic Data Processing)")
