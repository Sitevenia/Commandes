
import streamlit as st
import pandas as pd
import os
from datetime import datetime
import io
from modules.forecast import run_forecast_simulation, run_target_stock_sim

st.set_page_config(layout="wide", page_title="Forecast Hebdo")

st.title("Prévision des commandes hebdomadaires")

uploaded_file = st.file_uploader("Charger un fichier Excel", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.subheader("Données chargées")
    st.dataframe(df)

    st.subheader("Simulation standard")
    df_forecast = run_forecast_simulation(df)
    st.dataframe(df_forecast)

    if st.button("📤 Exporter la prévision standard en Excel"):
        output = io.BytesIO()
        df_forecast.to_excel(output, index=False, engine='openpyxl')
        st.download_button(
            label="📄 Télécharger la prévision standard",
            data=output.getvalue(),
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

        output2 = io.BytesIO()
        st.session_state["df_cible"].to_excel(output2, index=False, engine='openpyxl')
        st.download_button(
            label="📄 Télécharger la prévision cible",
            data=output2.getvalue(),
            file_name="prevision_cible.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Veuillez charger un fichier Excel pour démarrer.")
