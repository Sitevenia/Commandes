import streamlit as st
import pandas as pd
import os
from datetime import datetime
from modules.forecast import run_forecast_simulation, run_target_stock_simulation
from modules.rotation import detect_low_rotation_products
from modules.export import export_order_pdfs, export_low_rotation_list

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

    if st.button("Exporter les bons de commande PDF"):
        export_order_pdfs(df_forecast)
        st.success("Export PDF effectué dans output/pdf")

    st.subheader("Simulation par objectif de valeur de stock")
    objectif = st.number_input("Objectif de stock global (€)", min_value=0, step=100)
    if st.button("Lancer simulation cible"):
        df_cible = run_target_stock_sim(df, objectif)
        st.dataframe(df_cible)

    st.subheader("Produits à faible rotation")
    seuil = st.number_input("Seuil de rotation", min_value=0.0, value=10.0)
    if st.button("Extraire produits à faible rotation"):
        low_rotation = detect_low_rotation_products(df, threshold=seuil)
        export_low_rotation_list(low_rotation)
        st.dataframe(low_rotation)
        st.success("Liste exportée dans output/excel")
