
import streamlit as st
import pandas as pd
import os
from datetime import datetime
import io
import matplotlib.pyplot as plt
from modules.forecast import run_forecast_simulation, run_target_stock_sim

st.set_page_config(layout="wide", page_title="Forecast Hebdo")

st.title("Prévision des commandes hebdomadaires")

uploaded_file = st.file_uploader("Charger un fichier Excel", type=["xlsx"])

def format_excel(df, sheet_name):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
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
        valeur_finale = df_cible["Valeur totale"].sum()
        st.success(f"Simulation terminée. Valeur totale finale : {valeur_finale:,.2f} €")

    if "df_cible" in st.session_state:
        df_cible = st.session_state["df_cible"]
        st.dataframe(df_cible)

        excel_data = format_excel(df_cible, "Prévision cible")
        st.download_button(
            label="📄 Télécharger la prévision cible",
            data=excel_data.getvalue(),
            file_name="prevision_cible.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.subheader("📊 Répartition des quantités commandées")
        df_chart = df_cible[df_cible["Produit"] != "TOTAL"]
        fig, ax = plt.subplots(figsize=(10, 5))
        ax.bar(df_chart["Produit"], df_chart["Quantité commandée"], color="skyblue")
        ax.set_xlabel("Produit")
        ax.set_ylabel("Quantité commandée")
        ax.set_title("Répartition des quantités commandées par produit")
        plt.xticks(rotation=45, ha='right')
        st.pyplot(fig)
else:
    st.info("Veuillez charger un fichier Excel pour démarrer.")
