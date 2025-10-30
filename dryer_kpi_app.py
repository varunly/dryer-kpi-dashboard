import streamlit as st
import pandas as pd
import tempfile
import plotly.express as px
from dryer_kpi_monthly_final import main as run_kpi, CONFIG

# ------------------ Page Configuration ------------------
st.set_page_config(
    page_title="Lindner Dryer KPI Dashboard",
    page_icon="🏭", # Added a relevant icon
    layout="wide",
    initial_sidebar_state="expanded"
)

# ------------------ Custom CSS (UPDATED FOR REALISM) ------------------
st.markdown("""
    <style>
    /* Main background and font */
    body {
        background-color: #f5f7fa;
        color: #2b2b2b;
    }

    /* Title bar */
    .main-title {
        font-size: 36px;
        color: #003366;
        font-weight: 700;
        text-align: center;
        margin-bottom: 20px;
    }

    /* Subheadings */
    .section-header {
        color: #003366;
        font-size: 22px;
        font-weight: 600;
        margin-top: 40px;
        border-bottom: 2px solid #005691;
        padding-bottom: 6px;
    }

    /* KPI cards (UPDATED) */
    .metric-card {
        background-color: white;
        padding: 20px;
        border-radius: 10px;
        text-align: center;
        box-shadow: 0 4px 12px rgba(0,0,0,0.05); 
        border-top: 5px solid #005691; 
        transition: transform 0.2s;
        margin-bottom: 10px;
    }
    .metric-card:hover {
        transform: translateY(-3px);
    }
    .metric-card h3 {
        color: #555555;
        font-size: 16px;
        margin-bottom: 5px;
    }
    .metric-card h2 {
        color: #003366;
        font-size: 32px;
        margin: 0;
    }
    </style>
""", unsafe_allow_html=True)

# ------------------ Header ------------------
st.markdown('<div class="main-title">Lindner – Dryer KPI Monitoring Dashboard</div>', unsafe_allow_html=True)

st.write("Upload your **Energy** and **Hordenwagen** Excel files to visualize energy KPIs, trends, and efficiency by product and zone.")

# ------------------ Sidebar ------------------
with st.sidebar:
    st.image("https://www.karrieretag.org/wp-content/uploads/2023/10/lindner-logo-1.png", use_column_width=True)
    st.markdown("---")
    energy_file = st.file_uploader("📊 Upload Energy File (.xlsx)", type=["xlsx"])
    wagon_file = st.file_uploader("🚛 Upload Hordenwagen File (.xlsm, .xlsx)", type=["xlsm", "xlsx"])
    products = st.multiselect("🧱 Select Product(s):", ["L30","L32","L34","L36","L38","L40","N40","N44"], default=["L36"])
    month = st.number_input("📅 Month (1–12, 0 = all months):", 0, 12, 0)
    st.markdown("---")
    run_button = st.button("▶️ Run KPI Analysis")

# ------------------ Processing ------------------
if run_button:
    if not energy_file or not wagon_file:
        st.error("⚠️ Please upload both files before running analysis.")
    else:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_e, \
             tempfile.NamedTemporaryFile(delete=False, suffix=".xlsm") as tmp_w:
            tmp_e.write(energy_file.read())
            tmp_w.write(wagon_file.read())
            tmp_e.flush()
            tmp_w.flush()

            CONFIG["energy_file"] = tmp_e.name
            CONFIG["wagon_file"] = tmp_w.name
            CONFIG["product_filter"] = products if products else None
            CONFIG["month_filter"] = month if month != 0 else None
            CONFIG["output_file"] = "Dryer_KPI_WebApp_Results.xlsx"

            with st.spinner("⏳ Running KPI Analysis..."):
                run_kpi()

            # Load results
            summary = pd.read_excel(CONFIG["output_file"], sheet_name="Summary_By_Month_Zone")
            yearly = pd.read_excel(CONFIG["output_file"], sheet_name="Yearly_Summary")

            # --------------- KPI Cards ---------------
            st.markdown('<div class="section-header">📈 Summary KPIs</div>', unsafe_allow_html=True)

            total_energy = yearly["Energy_kWh"].sum()
            avg_kpi = yearly["kWh_per_m3"].mean()
            total_volume = yearly["Volume_m3"].sum()

            col1, col2, col3 = st.columns(3)
            col1.markdown(f'<div class="metric-card"><h3>Total Energy Consumed</h3><h2>{total_energy:,.0f} kWh</h2></div>', unsafe_allow_html=True)
            col2.markdown(f'<div class="metric-card"><h3>Avg. Energy Efficiency</h3><h2>{avg_kpi:,.2f} kWh/m³</h2></div>', unsafe_allow_html=True)
            col3.markdown(f'<div class="metric-card"><h3>Total Volume Processed</h3><h2>{total_volume:,.0f} m³</h2></div>', unsafe_allow_html=True)

            # --------------- Charts ---------------
            st.markdown('<div class="section-header">📊 KPI and Volume Analysis</div>', unsafe_allow_html=True)
            
            # Use two columns for better dashboard layout
            col_chart_1, col_chart_2 = st.columns([2, 1])

            # Chart 1: Monthly KPI (existing bar chart)
            with col_chart_1:
                st.markdown("### Monthly KPI (kWh/m³) by Zone")
                fig1 = px.bar(
                    summary,
                    x="Month", y="kWh_per_m3", color="Zone",
                    barmode="group",
                    hover_data=["Produkt", "Energy_kWh", "Volume_m3"],
                    color_discrete_sequence=px.colors.sequential.Blues_r,
                    title=""
                )
                fig1.update_layout(height=450, xaxis_title="Month", yaxis_title="kWh/m³", plot_bgcolor="white")
                st.plotly_chart(fig1, use_container_width=True)

            # Chart 2: Volume Breakdown (NEW Pie Chart)
            with col_chart_2:
                st.markdown("### Volume Breakdown ($\text{m}^3$) by Zone")
                fig3 = px.pie(
                    yearly,
                    values='Volume_m3',
                    names='Zone',
                    hole=.3, 
                    color_discrete_sequence=px.colors.qualitative.Set1,
                    title=""
                )
                fig3.update_traces(textinfo='percent+label', marker=dict(line=dict(color='#000000', width=1)))
                fig3.update_layout(height=450, margin=dict(t=50, b=0, l=0, r=0))
                st.plotly_chart(fig3, use_container_width=True)


            st.markdown('<div class="section-header">📉 Yearly KPI by Zone (Product Comparison)</div>', unsafe_allow_html=True)
            fig2 = px.bar(
                yearly,
                x="Zone", y="kWh_per_m3", color="Produkt",
                hover_data=["Energy_kWh", "Volume_m3"],
                text_auto=".2f",
                color_discrete_sequence=px.colors.qualitative.Set2
            )
            fig2.update_layout(height=450, xaxis_title="Zone", yaxis_title="kWh/m³", plot_bgcolor="white")
            st.plotly_chart(fig2, use_container_width=True)

            st.markdown('<div class="section-header">📁 Download Full Excel Report</div>', unsafe_allow_html=True)
            with open(CONFIG["output_file"], "rb") as f:
                st.download_button(
                    label="📥 Download Excel Results",
                    data=f.read(),
                    file_name="Dryer_KPI_Results.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            st.success("✅ KPI Analysis complete. Use the charts above to explore efficiency trends.")
