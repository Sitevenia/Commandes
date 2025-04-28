import streamlit as st
import pandas as pd
import numpy as np
import io

def calculer_quantite_a_commander(df, semaine_columns):
    """Calcule la quantité à commander en fonction des critères donnés."""
    # Calculer la moyenne des ventes sur la totalité des colonnes
    moyenne_totale = df[semaine_columns].mean(axis=1)

    # Calculer la moyenne des 12 dernières semaines
    moyenne_12_dernieres_semaines = df[semaine_columns[-12:]].mean(axis=1)

    # Calculer la moyenne des 12 semaines identiques en N-1
    moyenne_12_semaines_N1 = df[semaine_columns[:12]].mean(axis=1)

    # Appliquer la pondération
    quantite_ponderee = 0.7 * moyenne_12_dernieres_semaines + 0.3 * moyenne_12_semaines_N1

    # Calculer la quantité à commander pour les 3 prochaines semaines
    quantite_a_commander = (quantite_ponderee * 3) - df["Stock"]
    quantite_a_commander = quantite_a_commander.apply(lambda x: max(0, x))  # Ne pas commander des quantités négatives

    return quantite_a_commander

st.set_page_config(page_title="Forecast App", layout="wide")
st.title("📦 Application de Prévision des Commandes")

# Chargement du fichier principal
uploaded_file = st.file_uploader("📁 Charger le fichier Excel principal", type=["xlsx"])

if uploaded_file:
    try:
        # Lire le fichier Excel en utilisant la ligne 8 comme en-tête
        df = pd.read_excel(uploaded_file, sheet_name="Tableau final", header=7)
        st.success("✅ Fichier principal chargé avec succès.")

        # Utiliser la 9ème colonne comme point de départ
        start_index = 8  # Index 8 car les index commencent à 0
        semaine_columns = df.columns[start_index:].tolist()

        # Sélectionner toutes les colonnes numériques à partir de la 9ème colonne
        numeric_columns = df[semaine_columns].select_dtypes(include=[np.number]).columns.tolist()

        exclude_columns = ["Tarif d'achat", "Conditionnement", "Stock"]
        semaine_columns = [col for col in numeric_columns if col not in exclude_columns]

        for col in semaine_columns + exclude_columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

        # Calculer la quantité à commander
        df["Quantité à commander"] = calculer_quantite_a_commander(df, semaine_columns)

        st.subheader("Quantités à commander pour les 3 prochaines semaines")
        st.dataframe(df[["Référence fournisseur", "Référence produit", "Désignation", "Quantité à commander"]])

        # Export des quantités à commander
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df[["Référence fournisseur", "Référence produit", "Désignation", "Quantité à commander"]].to_excel(writer, sheet_name="Quantités_à_commander", index=False)
        output.seek(0)
        st.download_button("📥 Télécharger Quantités à commander", output, file_name="quantites_a_commander.xlsx")

    except Exception as e:
        st.error(f"❌ Erreur : {e}")
else:
    st.info("Veuillez charger le fichier principal pour commencer.")
