import streamlit as st
import pandas as pd
import numpy as np
import io

def calculer_quantite_a_commander(df, semaine_columns, montant_minimum, duree_semaines):
    # (La fonction reste inchangée)
    pass

st.set_page_config(page_title="Forecast App", layout="wide")
st.title("📦 Application de Prévision des Commandes")

# Chargement du fichier principal
uploaded_file = st.file_uploader("📁 Charger le fichier Excel principal", type=["xlsx"])

if uploaded_file:
    try:
        # Lire le fichier Excel en utilisant la ligne 8 comme en-tête
        df = pd.read_excel(uploaded_file, sheet_name="Tableau final", header=7)
        st.success("✅ Fichier principal chargé avec succès.")

        # Extraire la liste des fournisseurs
        fournisseurs = df["Nom du Fournisseur"].unique().tolist()  # Remplacez par le nom correct de la colonne

        # Widget pour sélectionner les fournisseurs
        selected_fournisseurs = st.multiselect(
            "Sélectionnez les fournisseurs",
            options=fournisseurs,
            default=fournisseurs  # Par défaut, tous les fournisseurs sont sélectionnés
        )

        # Filtrer les données en fonction des fournisseurs sélectionnés
        df_filtered = df[df["Nom du Fournisseur"].isin(selected_fournisseurs)]

        # Utiliser la colonne 13 comme point de départ
        start_index = 13  # Colonne "N"

        # Sélectionner toutes les colonnes numériques à partir de la colonne 13
        semaine_columns = df_filtered.columns[start_index:].tolist()
        numeric_columns = df_filtered[semaine_columns].select_dtypes(include=[np.number]).columns.tolist()

        exclude_columns = ["Tarif d'achat", "Conditionnement", "Stock"]
        semaine_columns = [col for col in numeric_columns if col not in exclude_columns]

        for col in semaine_columns + exclude_columns:
            df_filtered[col] = pd.to_numeric(df_filtered[col], errors="coerce").fillna(0)

        # Interface pour saisir la durée en semaines
        duree_semaines = st.number_input("Durée en semaines pour la commande", value=3, min_value=1, step=1)

        # Interface pour saisir le montant minimum de commande
        montant_minimum = st.number_input("Montant minimum de commande (€)", value=0.0, step=100.0)

        # Calculer la quantité à commander et les autres valeurs
        df_filtered["Quantité à commander"], df_filtered["Ventes N-1"], df_filtered["Ventes 12 semaines identiques N-1"], df_filtered["Ventes 12 dernières semaines"], montant_total = \
            calculer_quantite_a_commander(df_filtered, semaine_columns, montant_minimum, duree_semaines)

        # Ajouter la colonne "Tarif d'achat"
        df_filtered["Tarif d'achat"] = df_filtered["Tarif d'achat"]

        # Calculer la colonne "Total"
        df_filtered["Total"] = df_filtered["Tarif d'achat"] * df_filtered["Quantité à commander"]

        # Calculer la colonne "Stock à terme"
        df_filtered["Stock à terme"] = df_filtered["Stock"] + df_filtered["Quantité à commander"]

        # Vérifier si les colonnes nécessaires existent
        required_columns = ["AF_RefFourniss", "Référence Article", "Désignation Article", "Stock"]
        missing_columns = [col for col in required_columns if col not in df_filtered.columns]

        if missing_columns:
            st.error(f"❌ Colonnes manquantes dans le fichier : {missing_columns}")
        else:
            # Organiser l'ordre des colonnes pour l'affichage
            display_columns = required_columns + ["Ventes N-1", "Ventes 12 semaines identiques N-1", "Ventes 12 dernières semaines", "Conditionnement", "Quantité à commander", "Stock à terme", "Tarif d'achat", "Total"]

            # Afficher le montant total de la commande
            st.metric(label="Montant total de la commande", value=f"{montant_total:.2f} €")

            st.subheader("Quantités à commander pour les prochaines semaines")
            st.dataframe(df_filtered[display_columns])

            # Filtrer les produits pour lesquels il y a des quantités à commander pour l'exportation
            df_export = df_filtered[df_filtered["Quantité à commander"] > 0].copy()

            # Ajouter une ligne de total en bas du tableau filtré
            total_row = pd.DataFrame(df_export[["Total"]].sum()).T
            total_row.index = ["Total"]
            df_with_total = pd.concat([df_export[display_columns], total_row], ignore_index=False)

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
