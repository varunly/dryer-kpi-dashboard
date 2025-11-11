"""
Lindner Dryer - Production Order Optimizer
Standalone version with embedded optimizer
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

# ------------------ EMBEDDED OPTIMIZER CLASS ------------------
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

# ------------------ Load Database ------------------
@st.cache_resource
def load_database():
    """Load the optimization database"""
    
    # Try multiple paths
    possible_paths = [
        "optimization_database.json",
        "../optimization_database.json",
        "/mount/src/dryer-kpi-dashboard/optimization_database.json",
    ]
    
    for path in possible_paths:
        if os.path.exists(path):
            try:
                with open(path, 'r') as f:
                    db = json.load(f)
                print(f"✅ Loaded database from: {path}")
                return db, True
            except Exception as e:
                print(f"Error loading from {path}: {e}")
                continue
    
    return None, False

database, db_loaded = load_database()

# ------------------ Header ------------------
st.markdown('<div class="main-title">🔄 Lindner – Dryer Production Optimizer</div>', 
            unsafe_allow_html=True)

if not db_loaded:
    st.error("""
    ❌ **Optimization Database Not Found!**
    
    The file `optimization_database.json` must be in the repository root.
    
    **Current location checked:**
    - Root directory: `optimization_database.json`
    - Parent directory: `../optimization_database.json`
    - Absolute path: `/mount/src/dryer-kpi-dashboard/optimization_database.json`
    
    **Please ensure:**
    1. File exists at: https://github.com/Varunism/dryer-kpi-dashboard/optimization_database.json
    2. File is committed to the `main` branch
    3. File is in the ROOT directory (not in a subfolder)
    """)
    st.stop()

st.success("✅ Optimization database loaded successfully")

# Initialize optimizer
optimizer = ProductionOptimizer(database)

# ------------------ Sidebar ------------------
with st.sidebar:
    st.image("https://www.karrieretag.org/wp-content/uploads/2023/10/lindner-logo-1.png", 
             use_container_width=True)
    st.markdown("---")
    
    st.subheader("📦 Weekly Production Plan")
    st.write("Enter wagons needed per product:")
    
    all_products = list(database['product_profiles'].keys())
    
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

# ------------------ Main Content ------------------
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
                
                # Product details
                with st.expander("📊 Product Energy Profiles"):
                    product_details = []
                    for product in result['optimal_sequence']:
                        profile = optimizer.get_product_info(product)
                        product_details.append({
                            'Product': product,
                            'Type': profile['type'],
                            'Thickness (mm)': profile['thickness_mm'],
                            'kWh/m³': f"{profile['avg_kwh_per_m3']:.2f}",
                            'kWh/Wagon': f"{profile['kwh_per_wagon']:.1f}",
                            'Wagons': weekly_demand.get(product, 0),
                            'Total Energy': f"{profile['kwh_per_wagon'] * weekly_demand.get(product, 0):.0f} kWh"
                        })
                    
                    details_df = pd.DataFrame(product_details)
                    st.dataframe(details_df, use_container_width=True)
                
                # Export
                st.markdown("### 📥 Export Production Plan")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    # Excel export
                    excel_file = "Production_Plan.xlsx"
                    with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
                        seq_df = pd.DataFrame({
                            'Position': range(1, len(result['optimal_sequence']) + 1),
                            'Product': result['optimal_sequence']
                        })
                        seq_df.to_excel(writer, sheet_name='Sequence', index=False)
                        transitions_df.to_excel(writer, sheet_name='Transitions', index=False)
                        details_df.to_excel(writer, sheet_name='Product_Details', index=False)
                    
                    with open(excel_file, 'rb') as f:
                        st.download_button(
                            "📥 Download Excel Plan",
                            f.read(),
                            excel_file,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    
                    os.unlink(excel_file)
                
                with col2:
                    # Text report
                    report = f"""
LINDNER DRYER - PRODUCTION PLAN
Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}
{'='*50}

OPTIMAL SEQUENCE:
{' → '.join(result['optimal_sequence'])}

METRICS:
- Total Transition Cost: {result['total_transition_cost']:.1f} kWh
- Savings vs Worst Case: {result['savings_percent']:.1f}%
- Total Wagons: {sum(weekly_demand.values())}
- Products: {len(weekly_demand)}

WEEKLY DEMAND:
{chr(10).join([f'  {p}: {w} wagons' for p, w in weekly_demand.items()])}

RECOMMENDATIONS:
{chr(10).join([f'  • {rec}' for rec in result['recommendations']])}
                    """
                    
                    st.download_button(
                        "📄 Download Text Report",
                        report,
                        "production_plan.txt",
                        "text/plain",
                        use_container_width=True
                    )
                
                st.success("✅ Optimization complete!")

else:
    # Instructions
    st.markdown("""
    ## 🚀 How to Use This Optimizer
    
    ### 📦 Step 1: Enter Weekly Demand
    Use the sidebar to enter the number of wagons needed for each product type.
    
    ### 🎯 Step 2: Optimize
    Click the "Optimize Production Order" button.
    
    ### 📊 Step 3: Review Results
    You'll get:
    - **Optimal production sequence** to minimize energy
    - **Transition cost analysis** for each changeover
    - **Energy savings** compared to random ordering
    - **Specific recommendations** for your team
    - **Downloadable reports** (Excel and text)
    
    ---
    
    ## 💡 Why This Sequence is Optimal
    
    The optimizer considers:
    - ✅ Product thickness differences (setup time)
    - ✅ Material type changes (cleaning requirements)  
    - ✅ Energy consumption patterns
    - ✅ Temperature adjustment costs
    
    ---
    
    **Database Info:**
    - Products: {len(database['product_profiles'])}
    - Database created: {database['metadata'].get('created', 'Unknown')}
    """)

st.markdown("---")
st.caption("🏭 Lindner Dryer - Production Optimizer v2.0 (Standalone)")
