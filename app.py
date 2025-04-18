
import streamlit as st
import pandas as pd
import io
import matplotlib.pyplot as plt
from openpyxl.styles import PatternFill
from modules.forecast import run_forecast_simulation, run_target_stock_sim

st.set_page_config(layout="wide", page_title="Forecast Hebdo")

st.title("Prévision des commandes hebdomadaires")

uploaded_file = st.file_uploader("Charger un fichier Excel", type=["xlsx"])

def format_excel(df, sheet_name):
    output = io.BytesIO()
    df_export = df.copy()

    ventes_cols = [col for col in df_export.columns if col.startswith("2024-S")]
    export_order = [
        "Fournisseur", "Produit", "Désignation", "Stock", "Valeur stock actuel",
        "Conditionnement", "Tarif d’achat", "Quantité mini"
    ] + ventes_cols + [
        "Quantité commandée", "Stock total après commande", "Valeur ajoutée", "Valeur totale"
    ]

    
    # Appliquer l'ordre exact
    # Ligne TOTAL simplifiée uniquement sur colonnes numériques
    if "Produit" in df_export.columns and "TOTAL" not in df_export["Produit"].astype(str).str.upper().values:
        numeric_cols = df_export.select_dtypes(include=["number"]).columns
        total_data = {col: df_export[col].sum() if col in numeric_cols else "" for col in df_export.columns}
        total_data["Produit"] = "TOTAL"
        df_export.loc[len(df_export)] = total_data
    
    df_export = df_export[[col for col in export_order if col in df_export.columns]]

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name=sheet_name)
        worksheet = writer.sheets[sheet_name]

        euro_cols = ["Tarif d’achat", "Valeur stock actuel", "Valeur ajoutée", "Valeur totale"]
        for col_name in euro_cols:
            if col_name in df_export.columns:
                col_idx = df_export.columns.get_loc(col_name) + 1
                for row in range(2, len(df_export) + 2):
                    worksheet.cell(row=row, column=col_idx).number_format = u'€#,##0.00'

        last_row = len(df_export) + 1
        if "Produit" in df_export.columns and str(df_export.iloc[-1]["Produit"]).strip().upper() == "TOTAL":
            fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
            for col in range(1, len(df_export.columns) + 1):
                worksheet.cell(row=last_row, column=col).fill = fill

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
