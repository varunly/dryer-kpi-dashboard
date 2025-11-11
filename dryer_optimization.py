"""
Lindner Dryer - Production Order Optimizer
Uses pre-built database for instant optimization
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import json
import os

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

# ------------------ Load Optimizer ------------------
@st.cache_resource
def load_optimizer():
    """Load the optimization database"""
    try:
        from simple_optimizer import SimpleProductionOptimizer
        opt = SimpleProductionOptimizer("optimization_database.json")
        return opt, True
    except Exception as e:
        return None, False

optimizer, db_loaded = load_optimizer()

# ------------------ Header ------------------
st.markdown('<div class="main-title">🔄 Lindner Dryer - Production Optimizer</div>', 
            unsafe_allow_html=True)

if not db_loaded:
    st.error("""
    ❌ **Optimization Database Not Found!**
    
    Please run `build_optimization_database.py` first to create the database.
    
    ```bash
    python build_optimization_database.py
    ```
    
    This will analyze your historical data and create the optimization database.
    """)
    st.stop()

st.success("✅ Optimization database loaded successfully")

# ------------------ Sidebar ------------------
with st.sidebar:
    st.image("https://www.karrieretag.org/wp-content/uploads/2023/10/lindner-logo-1.png", 
             use_column_width=True)
    st.markdown("---")
    
    st.subheader("📦 Weekly Production Plan")
    
    st.write("Enter the number of wagons for each product:")
    
    all_products = ["L28", "L30", "L32", "L34", "L36", "L38", "L40", "L44", "N40", "N44", "U36"]
    
    weekly_demand = {}
    
    for product in all_products:
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
    optimize_button = st.button("🚀 Optimize Production Order", use_container_width=True)

# ------------------ Main Content ------------------
if optimize_button:
    if not weekly_demand:
        st.warning("⚠️ Please enter production quantities for at least one product")
    else:
        with st.spinner("🔄 Calculating optimal sequence..."):
            # Get products to optimize
            products_to_optimize = list(weekly_demand.keys())
            
            # Run optimization
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
                        delta=f"vs worst case",
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
                
                if result['recommendations']:
                    for rec in result['recommendations']:
                        st.info(rec)
                else:
                    st.success("✅ Optimal sequence with smooth transitions!")
                
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
                        # Sequence
                        seq_df = pd.DataFrame({
                            'Position': range(1, len(result['optimal_sequence']) + 1),
                            'Product': result['optimal_sequence']
                        })
                        seq_df.to_excel(writer, sheet_name='Sequence', index=False)
                        
                        # Transitions
                        transitions_df.to_excel(writer, sheet_name='Transitions', index=False)
                        
                        # Product details
                        details_df.to_excel(writer, sheet_name='Product_Details', index=False)
                    
                    with open(excel_file, 'rb') as f:
                        st.download_button(
                            "📥 Download Excel Plan",
                            f.read(),
                            excel_file,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                
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

TRANSITION DETAILS:
{chr(10).join([f'  {t["from"]} → {t["to"]}: {t["cost_kwh"]:.1f} kWh' for t in result['transitions']])}
                    """
                    
                    st.download_button(
                        "📄 Download Text Report",
                        report,
                        "production_plan.txt",
                        "text/plain",
                        use_container_width=True
                    )
                
                st.success("✅ Optimization complete! Use the sequence above for your production schedule.")

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
    - **Optimal production sequence** to minimize energy consumption
    - **Transition cost analysis** for each changeover
    - **Energy savings** compared to random ordering
    - **Specific recommendations** for your production team
    - **Downloadable reports** (Excel and text formats)
    
    ---
    
    ## 💡 Why This Sequence is Optimal
    
    The optimizer considers:
    - ✅ Actual energy consumption from your historical data
    - ✅ Product thickness differences (setup time)
    - ✅ Material type changes (cleaning requirements)
    - ✅ Temperature adjustment costs between products
    
    ### 🎓 Optimization Principles:
    
    1. **Group similar products** - Keeps same material types together
    2. **Gradual thickness changes** - Minimizes dryer adjustments
    3. **Energy-aware sequencing** - Considers actual kWh/m³ from data
    4. **Minimize type changes** - Reduces cleaning and quality risks
    
    ---
    
    ## 📈 Example
    
    **Weekly Demand:**
    - L36: 20 wagons
    - L38: 15 wagons
    - L30: 10 wagons
    - N40: 12 wagons
    
    **Optimal Sequence:**  
    `L30 → L36 → L38 → N40`
    
    **Why?**
    - Starts with thinnest L-type (L30)
    - Gradually increases thickness (L30→L36→L38)
    - Only one material type change (L→N) at the end
    - Saves ~25% energy vs random order
    
    ---
    
    **Need to update the database?**  
    Run `build_optimization_database.py` again with your latest data.
    """)
    
    # Show database info
    if optimizer:
        st.markdown("---")
        st.markdown("### 📊 Database Information")
        
        with open("optimization_database.json", 'r') as f:
            db = json.load(f)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("Products in Database", len(db['product_profiles']))
        
        with col2:
            created = datetime.fromisoformat(db['metadata']['created'])
            days_old = (datetime.now() - created).days
            st.metric("Database Age", f"{days_old} days")
        
        with col3:
            st.metric("Total Transitions", len(db['product_profiles'])**2)
        
        # Show available products
        with st.expander("View Available Products"):
            products_info = []
            for prod, profile in db['product_profiles'].items():
                products_info.append({
                    'Product': prod,
                    'Type': profile['type'],
                    'Thickness': f"{profile['thickness_mm']}mm",
                    'kWh/m³': f"{profile['avg_kwh_per_m3']:.2f}",
                    'Data Points': profile['total_wagons_produced']
                })
            
            st.dataframe(pd.DataFrame(products_info), use_container_width=True)
