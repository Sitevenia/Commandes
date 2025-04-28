import streamlit as st
import pandas as pd
import numpy as np
import io

def calculer_quantite_a_commander(df, semaine_columns, montant_minimum):
    """Calcule la quantité à commander en fonction des critères donnés."""
    # Calculer la moyenne des ventes sur la totalité des colonnes (Ventes N-1)
    ventes_N1 = df[semaine_columns].sum(axis=1)

    # Calculer la moyenne des 12 dernières semaines
    ventes_12_dernieres_semaines = df[semaine_columns[-12:]].sum(axis=1)

    # Calculer la moyenne des 12 semaines identiques en N-1
    ventes_12_semaines_N1 = df[semaine_columns[-64:-52]].sum(axis=1)

    # Appliquer la pondération
    quantite_ponderee = 0.7 * (ventes_12_dernieres_semaines / 12) + 0.3 * (ventes_12_semaines_N1 / 12)

    # Calculer la quantité à commander pour les 3 prochaines semaines
    quantite_a_commander = (quantite_ponderee * 3) - df["Stock"]
    quantite_a_commander = quantite_a_commander.apply(lambda x: max(0, x))  # Ne pas commander des quantités négatives

    # Ajuster les quantités à commander pour qu'elles soient des multiples entiers des conditionnements
    conditionnement = df["Conditionnement"]
    quantite_a_commander = [int(np.ceil(q / cond) * cond) for q, cond in zip(quantite_a_commander, conditionnement)]

    # Calculer le montant total initial
    montant_total_initial = (df["Tarif d'achat"] * quantite_a_commander).sum()

    # Si le montant minimum est supérieur au montant calculé, ajuster les quantités
    if montant_minimum > 0 and montant_minimum > montant_total_initial:
        for i in range(len(quantite_a_commander)):
            while montant_total_initial < montant_minimum:
                quantite_a_commander[i] += conditionnement[i]
                montant_total_initial = (df["Tarif d'achat"] * quantite_a_commander).sum()
                if montant_total_initial >= montant_minimum:
                    break

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

        # Afficher les noms des colonnes pour vérification
        st.write("Noms des colonnes :", df.columns.tolist())

        # Utiliser la colonne "202401" comme point de départ
        start_column = "202401"
        if start_column in df.columns:
            start_index = df.columns.get_loc(start_column)
        else:
            st.error(f"❌ Colonne '{start_column}' non trouvée dans le fichier.")
            start_index = None

        if start_index is not None:
            # Sélectionner toutes les colonnes numériques à partir de "202401"
            semaine_columns = df.columns[start_index:].tolist()
            numeric_columns = df[semaine_columns].select_dtypes(include=[np.number]).columns.tolist()

            exclude_columns = ["Tarif d'achat", "Conditionnement", "Stock"]
            semaine_columns = [col for col in numeric_columns if col not in exclude_columns]

            for col in semaine_columns + exclude_columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

            # Interface pour saisir le montant minimum de commande
            montant_minimum = st.number_input("Montant minimum de commande (€)", value=0.0, step=100.0)

            # Calculer la quantité à commander et les autres valeurs
            df["Quantité à commander"], df["Ventes N-1"], df["Ventes 12 semaines identiques N-1"], df["Ventes 12 dernières semaines"] = \
                calculer_quantite_a_commander(df, semaine_columns, montant_minimum)

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
        else:
            st.error("❌ Impossible de trouver la colonne de départ.")

    except Exception as e:
        st.error(f"❌ Erreur : {e}")
else:
    st.info("Veuillez charger le fichier principal pour commencer.")
