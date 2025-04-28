import streamlit as st
import pandas as pd
import numpy as np
import io

def ajustement_commandes_exceptionnelles(df, semaine_columns):
    """Ajuste les valeurs exceptionnelles dans les ventes."""
    # Calculer la moyenne des ventes hebdomadaires sur l'ensemble des colonnes
    moyenne_hebdo = df[semaine_columns].mean(axis=1)

    # Remplacer les valeurs exceptionnelles par la moyenne hebdomadaire
    df_ajuste = df[semaine_columns].apply(lambda row: row.apply(lambda x: moyenne_hebdo if x > 3 * moyenne_hebdo else x), axis=1)

    return df_ajuste

def calculer_quantite_a_commander(df, semaine_columns):
    """Calcule la quantité à commander en fonction des critères donnés."""
    # Ajuster les valeurs exceptionnelles
    df_ajuste = ajustement_commandes_exceptionnelles(df, semaine_columns)

    # Calculer la moyenne des ventes sur la totalité des colonnes (Ventes N-1)
    ventes_N1 = df_ajuste.sum(axis=1)

    # Calculer la moyenne des 12 dernières semaines
    ventes_12_dernieres_semaines = df_ajuste[semaine_columns[-12:]].sum(axis=1)

    # Calculer la moyenne des 12 semaines identiques en N-1
    ventes_12_semaines_N1 = df_ajuste[semaine_columns[-64:-52]].sum(axis=1)

    # Appliquer la pondération
    quantite_ponderee = 0.7 * (ventes_12_dernieres_semaines / 12) + 0.3 * (ventes_12_semaines_N1 / 12)

    # Calculer la quantité à commander pour les 3 prochaines semaines
    quantite_a_commander = (quantite_ponderee * 3) - df["Stock"]
    quantite_a_commander = quantite_a_commander.apply(lambda x: max(0, x))  # Ne pas commander des quantités négatives

    # Ajuster les quantités à commander pour qu'elles soient des multiples entiers des conditionnements
    conditionnement = df["Conditionnement"]
    quantite_a_commander = [int(np.ceil(q / cond) * cond) for q, cond in zip(quantite_a_commander, conditionnement)]

    return quantite_a_commander, ventes_N1, ventes_12_semaines_N1, ventes_12_dernieres_semaines

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

        # Calculer la quantité à commander et les autres valeurs
        df["Quantité à commander"], df["Ventes N-1"], df["Ventes 12 semaines identiques N-1"], df["Ventes 12 dernières semaines"] = \
            calculer_quantite_a_commander(df, semaine_columns)

        # Ajouter la colonne "Tarif d'achat"
        df["Tarif d'achat"] = df["Tarif d'achat"]

        # Calculer la colonne "Total"
        df["Total"] = df["Tarif d'achat"] * df["Quantité à commander"]

        # Vérifier si les colonnes nécessaires existent
        required_columns = ["AF_RefFourniss", "Référence Article", "Désignation Article", "Stock"]
        missing_columns = [col for col in required_columns if col not in df.columns]

        if missing_columns:
            st.error(f"❌ Colonnes manquantes dans le fichier : {missing_columns}")
        else:
            # Organiser l'ordre des colonnes pour l'affichage et l'exportation
            display_columns = required_columns + ["Ventes N-1", "Ventes 12 semaines identiques N-1", "Ventes 12 dernières semaines", "Conditionnement", "Quantité à commander", "Tarif d'achat", "Total"]

            # Ajouter une ligne de total en bas du tableau
            total_row = pd.DataFrame(df[["Total"]].sum()).T
            total_row.index = ["Total"]
            df_with_total = pd.concat([df[display_columns], total_row], ignore_index=False)

            st.subheader("Quantités à commander pour les 3 prochaines semaines")
            st.dataframe(df_with_total)

            # Export des quantités à commander
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_with_total.to_excel(writer, sheet_name="Quantités_à_commander", index=False)
            output.seek(0)
            st.download_button("📥 Télécharger Quantités à commander", output, file_name="quantites_a_commander.xlsx")

    except Exception as e:
        st.error(f"❌ Erreur : {e}")
else:
    st.info("Veuillez charger le fichier principal pour commencer.")
