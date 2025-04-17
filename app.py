
import streamlit as st
import pandas as pd
import os
from datetime import datetime
import io
from modules.forecast import run_forecast_simulation, run_target_stock_sim

st.set_page_config(layout="wide", page_title="Forecast Hebdo")

st.title("Prévision des commandes hebdomadaires")

uploaded_file = st.file_uploader("Charger un fichier Excel", type=["xlsx"])

def format_excel(df, sheet_name):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)

        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        currency_cols = ["Tarif d’achat", "Valeur stock actuel", "Valeur ajoutée", "Valeur totale"]
        for col_name in currency_cols:
            if col_name in df.columns:
                col_idx = df.columns.get_loc(col_name) + 1
                col_letter = chr(64 + col_idx) if col_idx <= 26 else f"A{chr(64 + col_idx - 26)}"
                for row in range(2, len(df) + 2):
                    worksheet[f"{col_letter}{row}"].number_format = "€#,##0.00"

    return output

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.subheader("Données chargées")
    st.dataframe(df)

    st.subheader("Simulation standard")
    df_forecast = run_forecast_simulation(df)
    st.dataframe(df_forecast)

    if st.button("📤 Exporter la prévision standard en Excel"):
        excel_data = format_excel(df_forecast, "Prévision standard")
        st.download_button(
            label="📄 Télécharger la prévision standard",
            data=excel_data.getvalue(),
            file_name="prevision_standard.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.subheader("Simulation par objectif de valeur de stock")
    objectif = st.number_input("Objectif de stock global (€)", min_value=0, step=100)

    if st.button("Lancer simulation cible"):
        df_cible = run_target_stock_sim(df, objectif)
        st.session_state["df_cible"] = df_cible
        st.success("Simulation cible générée.")

    if "df_cible" in st.session_state:
        st.dataframe(st.session_state["df_cible"])

        excel_data = format_excel(st.session_state["df_cible"], "Prévision cible")
        st.download_button(
            label="📄 Télécharger la prévision cible",
            data=excel_data.getvalue(),
            file_name="prevision_cible.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Veuillez charger un fichier Excel pour démarrer.")
