import streamlit as st
import pandas as pd
import tempfile
import plotly.express as px
from dryer_kpi_monthly_final import main as run_kpi, CONFIG

st.set_page_config(page_title="Dryer KPI Dashboard", layout="wide")

st.title("📊 Dryer KPI Dashboard (Lindner)")

st.markdown("""
Upload your **Energy Data** and **Hordenwagen File**, choose your product and month,  
and see the KPI results (kWh/m³) directly here with interactive charts.
""")

energy_file = st.file_uploader("Upload Energy File (.xlsx)", type=["xlsx"])
wagon_file = st.file_uploader("Upload Hordenwagen File (.xlsm, .xlsx)", type=["xlsm", "xlsx"])

product = st.text_input("Enter Product (e.g. L36, L38, N40) or leave blank for all:", "")
month = st.number_input("Month (1–12, 0 = all months):", 0, 12, 0)

if st.button("▶️ Run Analysis"):
    if not energy_file or not wagon_file:
        st.error("Please upload both files first.")
    else:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_e, \
             tempfile.NamedTemporaryFile(delete=False, suffix=".xlsm") as tmp_w:
            tmp_e.write(energy_file.read())
            tmp_w.write(wagon_file.read())
            tmp_e.flush()
            tmp_w.flush()

            CONFIG["energy_file"] = tmp_e.name
            CONFIG["wagon_file"] = tmp_w.name
            CONFIG["product_filter"] = [product] if product else None
            CONFIG["month_filter"] = month if month != 0 else None
            CONFIG["output_file"] = "Dryer_KPI_WebApp_Results.xlsx"

            with st.spinner("Calculating KPIs... Please wait ⏳"):
                run_kpi()

            st.success("✅ KPI analysis completed successfully!")

            # Load the summary sheets
            summary = pd.read_excel(CONFIG["output_file"], sheet_name="Summary_By_Month_Zone")
            yearly = pd.read_excel(CONFIG["output_file"], sheet_name="Yearly_Summary")

            st.header("📈 Monthly KPI (kWh/m³)")
            fig1 = px.bar(
                summary, 
                x="Month", y="kWh_per_m3", color="Zone",
                hover_data=["Produkt"],
                barmode="group", 
                title="Monthly Energy KPI per Zone (kWh/m³)"
            )
            st.plotly_chart(fig1, use_container_width=True)

            st.header("📊 Yearly KPI Summary")
            fig2 = px.bar(
                yearly,
                x="Zone", y="kWh_per_m3", color="Produkt",
                hover_data=["Energy_kWh", "Volume_m3"],
                text_auto=".2f",
                title="Yearly KPI per Product and Zone"
            )
            st.plotly_chart(fig2, use_container_width=True)

            st.header("📁 Download Full Excel Report")
            st.download_button(
                label="📥 Download Excel Results",
                data=open(CONFIG["output_file"], "rb").read(),
                file_name="Dryer_KPI_WebApp_Results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
