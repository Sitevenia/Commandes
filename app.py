import streamlit as st
import pandas as pd
import numpy as np
import io

def calculer_quantite_a_commander(df, semaine_columns, montant_minimum, duree_semaines):
    """Calcule la quantité à commander en fonction des critères donnés."""
    # Calculer la moyenne des ventes sur la totalité des colonnes (Ventes N-1)
    ventes_N1 = df[semaine_columns].sum(axis=1)

    # Calculer la moyenne des 12 dernières semaines
    ventes_12_dernieres_semaines = df[semaine_columns[-12:]].sum(axis=1)

    # Calculer la moyenne des 12 semaines identiques en N-1
    ventes_12_semaines_N1 = df[semaine_columns[-64:-52]].sum(axis=1)

    # Calculer la moyenne des 12 semaines suivantes en N-1
    ventes_12_semaines_N1_suivantes = df[semaine_columns[-52:-40]].sum(axis=1)

    # Appliquer la pondération
    quantite_ponderee = 0.5 * (ventes_12_dernieres_semaines / 12) + 0.2 * (ventes_12_semaines_N1 / 12) + 0.3 * (ventes_12_semaines_N1_suivantes / 12)

    # Calculer la quantité nécessaire pour couvrir les ventes pendant la durée spécifiée
    quantite_necessaire = quantite_ponderee * duree_semaines

    # Calculer la quantité à commander
    quantite_a_commander = quantite_necessaire - df["Stock"]
    quantite_a_commander = quantite_a_commander.apply(lambda x: max(0, x))  # Ne pas commander des quantités négatives

    # Ajuster les quantités à commander pour qu'elles soient des multiples entiers des conditionnements
    conditionnement = df["Conditionnement"]
    quantite_a_commander = [int(np.ceil(q / cond) * cond) if q > 0 else 0 for q, cond in zip(quantite_a_commander, conditionnement)]

    # Vérifier si un produit est vendu au moins deux fois sur les 12 dernières semaines et si le stock est inférieur ou égal à 1
    for i in range(len(quantite_a_commander)):
        ventes_recentes = df[semaine_columns[-12:]].iloc[i]
        if (ventes_recentes > 0).sum() >= 2 and df["Stock"].iloc[i] <= 1:
            quantite_a_commander[i] = max(quantite_a_commander[i], conditionnement[i])

    # Vérifier si les quantités vendues en N-1 sont inférieures à 6 et si les ventes des 12 dernières semaines sont inférieures à 2
    for i in range(len(quantite_a_commander)):
        if ventes_N1.iloc[i] < 6 and ventes_12_dernieres_semaines.iloc[i] < 2:
            quantite_a_commander[i] = 0

    # Calculer le montant total initial
    montant_total_initial = (df["Tarif d'achat"] * quantite_a_commander).sum()

    # Si le montant minimum est supérieur au montant calculé, ajuster les quantités
    if montant_minimum > 0 and montant_total_initial < montant_minimum:
        while montant_total_initial < montant_minimum:
            for i in range(len(quantite_a_commander)):
                if quantite_a_commander[i] > 0:  # Augmenter seulement si une quantité est déjà commandée
                    quantite_a_commander[i] += conditionnement[i]
                    montant_total_initial = (df["Tarif d'achat"] * quantite_a_commander).sum()
                    if montant_total_initial >= montant_minimum:
                        break

    return quantite_a_commander, ventes_N1, ventes_12_semaines_N1, ventes_12_dernieres_semaines, montant_total_initial

def generer_rapport_excel(df, df_fournisseurs, montant_total):
    """Génère un rapport Excel avec les quantités à commander et le minimum de commande."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Écrire les quantités à commander
        df_with_total = pd.concat([df, pd.DataFrame([["Total", "", "", "", "", "", "", montant_total]], columns=df.columns + ["Total"])], ignore_index=True)

        # Ajouter une ligne pour le minimum de commande
        min_commande_info = df_fournisseurs.iloc[0]  # Supposons que nous prenons le premier fournisseur pour l'exemple
        min_commande_text = f"Minimum de commande : {min_commande_info['Montant minimum de commande']} €"
        min_commande_row = pd.DataFrame([["", "", "", "", "", "", min_commande_text, ""]], columns=df.columns + ["Total"])

        # Concaténer la ligne de minimum de commande avec les données
        df_with_total = pd.concat([min_commande_row, df_with_total], ignore_index=True)

        df_with_total.to_excel(writer, sheet_name="Quantités_à_commander", index=False)

    output.seek(0)
    return output

st.set_page_config(page_title="Forecast App", layout="wide")
st.title("📦 Application de Prévision des Commandes")

# Chargement du fichier principal
uploaded_file = st.file_uploader("📁 Charger le fichier Excel principal", type=["xlsx"])

if uploaded_file:
    try:
        # Lire le fichier Excel en utilisant la ligne 8 comme en-tête
        df = pd.read_excel(uploaded_file, sheet_name="Tableau final", header=7)
        st.success("✅ Fichier principal chargé avec succès.")

        # Lire l'onglet "Minimum de commande"
        df_fournisseurs = pd.read_excel(uploaded_file, sheet_name="Minimum de commande")

        # Utiliser la colonne 13 comme point de départ
        start_index = 13  # Colonne "N"

        # Sélectionner toutes les colonnes numériques à partir de la colonne 13
        semaine_columns = df.columns[start_index:].tolist()
        numeric_columns = df[semaine_columns].select_dtypes(include=[np.number]).columns.tolist()

        exclude_columns = ["Tarif d'achat", "Conditionnement", "Stock"]
        semaine_columns = [col for col in numeric_columns if col not in exclude_columns]

        for col in semaine_columns + exclude_columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

        # Interface pour saisir la durée en semaines
        duree_semaines = st.number_input("Durée en semaines pour la commande", value=3, min_value=1, step=1)

        # Interface pour saisir le montant minimum de commande
        montant_minimum = st.number_input("Montant minimum de commande (€)", value=0.0, step=100.0)

        # Calculer la quantité à commander et les autres valeurs
        df["Quantité à commander"], df["Ventes N-1"], df["Ventes 12 semaines identiques N-1"], df["Ventes 12 dernières semaines"], montant_total = \
            calculer_quantite_a_commander(df, semaine_columns, montant_minimum, duree_semaines)

        # Ajouter la colonne "Tarif d'achat"
        df["Tarif d'achat"] = df["Tarif d'achat"]

        # Calculer la colonne "Total"
        df["Total"] = df["Tarif d'achat"] * df["Quantité à commander"]

        # Calculer la colonne "Stock à terme"
        df["Stock à terme"] = df["Stock"] + df["Quantité à commander"]

        # Vérifier si les colonnes nécessaires existent
        required_columns = ["AF_RefFourniss", "Référence Article", "Désignation Article", "Stock"]
        missing_columns = [col for col in required_columns if col not in df.columns]

        if missing_columns:
            st.error(f"❌ Colonnes manquantes dans le fichier : {missing_columns}")
        else:
            # Organiser l'ordre des colonnes pour l'affichage
            display_columns = required_columns + ["Ventes N-1", "Ventes 12 semaines identiques N-1", "Ventes 12 dernières semaines", "Conditionnement", "Quantité à commander", "Stock à terme", "Tarif d'achat", "Total"]

            # Filtrer les produits pour lesquels il y a des quantités à commander pour l'exportation
            df_filtered = df[df["Quantité à commander"] > 0].copy()

            # Générer le rapport Excel
            output = generer_rapport_excel(df_filtered[display_columns], df_fournisseurs, montant_total)

            # Export des quantités à commander
            st.download_button("📥 Télécharger Quantités à commander", output, file_name="quantites_a_commander.xlsx")

    except Exception as e:
        st.error(f"❌ Erreur : {e}")
else:
    st.info("Veuillez charger le fichier principal pour commencer.")
