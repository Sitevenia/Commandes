
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

    # Bouton d'export Excel
    if st.button("📤 Exporter la prévision en Excel"):
        output = io.BytesIO()
        df_forecast.to_excel(output, index=False, engine='openpyxl')
        st.download_button(
            label="Télécharger le fichier Excel",
            data=output.getvalue(),
            file_name="prevision_commandes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.subheader("Simulation par objectif de valeur de stock")
    objectif = st.number_input("Objectif de stock global (€)", min_value=0, step=100)
    if st.button("Lancer simulation cible"):
        df_cible = run_target_stock_sim(df, objectif)
        st.dataframe(df_cible)
else:
    st.info("Veuillez charger un fichier Excel pour démarrer.")
