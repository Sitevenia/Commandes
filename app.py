
import streamlit as st
import pandas as pd
import io
from modules.forecast import run_forecast_simulation, run_target_stock_sim

def export_three_sheets(df_standard, df_target):
    output = io.BytesIO()

    df_comp = df_standard.copy()
    df_target_renamed = df_target.copy()

    df_comp = df_comp.rename(columns={
        "Fournisseur": "Fournisseur (standard)",
        "Stock": "Stock (standard)",
        "Conditionnement": "Conditionnement (standard)",
        "Tarif d’achat": "Tarif d’achat (standard)",
        "Quantité mini": "Quantité mini (standard)",
        "Valeur stock actuel": "Valeur stock actuel (standard)",
        "Quantité commandée": "Quantité commandée (standard)",
        "Valeur ajoutée": "Valeur ajoutée (standard)",
        "Valeur totale": "Valeur totale (standard)"
    })

    df_target_renamed = df_target_renamed.rename(columns={
        "Quantité commandée": "Quantité commandée (objectif)",
        "Valeur ajoutée": "Valeur ajoutée (objectif)",
        "Valeur totale": "Valeur totale (objectif)",
        "Stock total après commande": "Stock total après commande (objectif)"
    })

    df_comparatif = pd.merge(
        df_comp,
        df_target_renamed[[
            "Produit", "Désignation",
            "Quantité commandée (objectif)",
            "Valeur ajoutée (objectif)",
            "Valeur totale (objectif)",
            "Stock total après commande (objectif)"
        ]],
        on=["Produit", "Désignation"],
        how="outer"
    )

    export_columns = [
        "Produit", "Désignation", "Fournisseur (standard)", "Stock (standard)", "Conditionnement (standard)",
        "Tarif d’achat (standard)", "Quantité mini (standard)", "Valeur stock actuel (standard)",
        "Quantité commandée (standard)", "Valeur ajoutée (standard)", "Valeur totale (standard)",
        "Quantité commandée (objectif)", "Valeur ajoutée (objectif)", "Valeur totale (objectif)",
        "Stock total après commande (objectif)"
    ]

    df_comparatif = df_comparatif[[col for col in export_columns if col in df_comparatif.columns]]

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_standard.to_excel(writer, sheet_name="Simulation standard", index=False)
        df_target.to_excel(writer, sheet_name="Simulation objectif", index=False)
        df_comparatif.to_excel(writer, sheet_name="Comparatif", index=False)

    return output

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

    st.subheader("Simulation par objectif de valeur de stock")
    objectif = st.number_input("Objectif de stock global (€)", min_value=0, step=100)

    if st.button("Lancer simulation cible"):
        df_cible = run_target_stock_sim(df, objectif)
        st.session_state["df_cible"] = df_cible
        st.session_state["df_standard"] = df_forecast
        st.success(f"Simulation terminée avec objectif de {objectif:,.2f} €")

    if "df_cible" in st.session_state and "df_standard" in st.session_state:
        df_cible = st.session_state["df_cible"]
        df_forecast = st.session_state["df_standard"]

        st.subheader("Résultat simulation avec objectif")
        st.dataframe(df_cible)

        if st.button("📤 Exporter les 3 onglets Excel"):
            export_excel = export_three_sheets(df_forecast, df_cible)
            st.download_button(
                label="📄 Télécharger les 3 simulations (XLSX)",
                data=export_excel.getvalue(),
                file_name="export_forecast_complet.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.info("Veuillez charger un fichier Excel pour démarrer.")
