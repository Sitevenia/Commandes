import streamlit as st
import pandas as pd
import numpy as np
import io

def repartir_et_ajuster(total_qte, ventes_n1_semaines, conditionnement):
    """Répartit une quantité totale selon la saisonnalité et ajuste aux conditionnements."""
    if total_qte <= 0 or conditionnement is None or conditionnement <= 0:
        return [0] * len(ventes_n1_semaines)

    total_ventes = sum(ventes_n1_semaines)
    if total_ventes == 0:
        repartition = [1] * len(ventes_n1_semaines)
        total_ventes = len(ventes_n1_semaines)
    else:
        repartition = ventes_n1_semaines

    proportions = [v / total_ventes for v in repartition]
    qtes_hebdomadaires = [round(total_qte * p) for p in proportions]

    # Ajuster chaque semaine au multiple de conditionnement
    qtes_conditionnees = [int(round(q / conditionnement)) * conditionnement for q in qtes_hebdomadaires]

    # Réajuster si trop ou pas assez
    ecart = sum(qtes_conditionnees) - total_qte
    while ecart != 0:
        for i in range(len(qtes_conditionnees)):
            if ecart > 0 and qtes_conditionnees[i] >= conditionnement:
                qtes_conditionnees[i] -= conditionnement
                ecart -= conditionnement
            elif ecart < 0:
                qtes_conditionnees[i] += conditionnement
                ecart += conditionnement
            if ecart == 0:
                break

    return qtes_conditionnees

st.set_page_config(page_title="Forecast App", layout="wide")
st.title("📦 Application de Prévision des Commandes")

def safe_int(x):
    try:
        if pd.isna(x) or not np.isfinite(x):
            return 0
        return int(round(x))
    except:
        return 0

# Chargement du fichier principal
uploaded_file = st.file_uploader("📁 Charger le fichier Excel principal", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, sheet_name="Tableau final")
        st.success("✅ Fichier principal chargé avec succès.")

        # Trouver l'index de la colonne "202401"
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

            exclude_columns = ["Tarif d'achat", "Conditionnement"]
            semaine_columns = [col for col in numeric_columns if col not in exclude_columns]

            for col in semaine_columns + exclude_columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

            df["Conditionnement"] = df["Conditionnement"].replace(0, 1)
            df["Total ventes N-1"] = df[semaine_columns].sum(axis=1).replace(0, np.nan)
            saisonnalite = df[semaine_columns].div(df["Total ventes N-1"], axis=0).replace([np.inf, -np.inf], 0).fillna(0)

            # Sélection du type de simulation
            simulation_type = st.selectbox("Sélectionnez le type de simulation", ["Simulation simple", "Simulation avec objectif de montant"])

            if simulation_type == "Simulation simple":
                st.subheader("Simulation simple : progression personnalisée")
                progression = st.number_input("📈 Progression (%)", value=0.0, step=1.0)
                df["Qté Sim 1"] = df["Total ventes N-1"] * (1 + progression / 100)
                df["Qté Sim 1"] = (np.ceil(df["Qté Sim 1"] / df["Conditionnement"]) * df["Conditionnement"]).fillna(0).astype(int)

                if st.button("▶️ Lancer la Simulation simple"):
                    for i in df.index:
                        repartition = repartir_et_ajuster(
                            df.at[i, "Qté Sim 1"],
                            saisonnalite.loc[i, semaine_columns],
                            df.at[i, "Conditionnement"]
                        )
                        # Assurez-vous que la longueur de repartition correspond à celle des colonnes sélectionnées
                        if len(repartition) == len(semaine_columns):
                            df.loc[i, semaine_columns] = repartition
                        else:
                            st.error("Erreur : La longueur de la répartition ne correspond pas aux semaines sélectionnées.")

                    df["Montant Sim 1"] = df["Qté Sim 1"] * df["Tarif d'achat"]
                    total_sim1 = df["Montant Sim 1"].sum()
                    st.metric("💰 Total Simulation simple", f"€ {total_sim1:,.2f}")

                    # Export Simulation simple
                    output1 = io.BytesIO()
                    with pd.ExcelWriter(output1, engine="xlsxwriter") as writer:
                        # Filtrer les colonnes avant l'exportation
                        df_filtered = df[["Référence fournisseur", "Référence produit", "Désignation", "Qté Sim 1", "Montant Sim 1"] + semaine_columns]
                        df_filtered.to_excel(writer, sheet_name="Simulation_simple", index=False)
                    output1.seek(0)
                    st.download_button("📥 Télécharger Simulation simple", output1, file_name="simulation_simple.xlsx")

            elif simulation_type == "Simulation avec objectif de montant":
                st.subheader("Simulation avec objectif de montant")
                objectif = st.number_input("🎯 Objectif (€)", value=0.0, step=1000.0)

                if objectif > 0:
                    if st.button("▶️ Lancer la Simulation avec objectif de montant"):
                        df_sim2 = df.copy()
                        df_sim2["Qté Base"] = df["Total ventes N-1"].replace(0, 1)
                        total_base_value = (df_sim2["Qté Base"] * df_sim2["Tarif d'achat"]).sum()

                        best_coef = 1.0
                        best_diff = float("inf")
                        for coef in np.arange(0.01, 2.0, 0.01):
                            q_test = np.ceil((df_sim2["Qté Base"] * coef) / df_sim2["Conditionnement"]) * df_sim2["Conditionnement"]
                            montant_test = (q_test * df_sim2["Tarif d'achat"]).sum()
                            diff = abs(montant_test - objectif)
                            if montant_test <= objectif and diff < best_diff:
                                best_diff = diff
                                best_coef = coef

                        df_sim2["Qté Sim 2"] = (np.ceil((df_sim2["Qté Base"] * best_coef) / df_sim2["Conditionnement"]) * df_sim2["Conditionnement"]).fillna(0).astype(int)

                        for i in df_sim2.index:
                            repartition = repartir_et_ajuster(
                                df_sim2.at[i, "Qté Sim 2"],
                                saisonnalite.loc[i, semaine_columns],
                                df_sim2.at[i, "Conditionnement"]
                            )
                            # Assurez-vous que la longueur de repartition correspond à celle des colonnes sélectionnées
                            if len(repartition) == len(semaine_columns):
                                df_sim2.loc[i, semaine_columns] = repartition
                            else:
                                st.error("Erreur : La longueur de la répartition ne correspond pas aux semaines sélectionnées.")

                        df_sim2["Montant Sim 2"] = df_sim2["Qté Sim 2"] * df_sim2["Tarif d'achat"]
                        total_sim2 = df_sim2["Montant Sim 2"].sum()
                        st.metric("✅ Montant Simulation avec objectif de montant", f"€ {total_sim2:,.2f}")

                        st.dataframe(df_sim2[["Référence fournisseur", "Référence produit", "Désignation", "Qté Sim 2", "Montant Sim 2"]])

                        # Export Simulation avec objectif de montant
                        output2 = io.BytesIO()
                        with pd.ExcelWriter(output2, engine="xlsxwriter") as writer:
                            # Filtrer les colonnes avant l'exportation
                            df_filtered_sim2 = df_sim2[["Référence fournisseur", "Référence produit", "Désignation", "Qté Sim 2", "Montant Sim 2"] + semaine_columns]
                            df_filtered_sim2.to_excel(writer, sheet_name="Simulation_objectif", index=False)
                        output2.seek(0)
                        st.download_button("📥 Télécharger Simulation avec objectif de montant", output2, file_name="simulation_objectif.xlsx")
        else:
            st.error("❌ Impossible de trouver la colonne de départ.")

    except Exception as e:
        st.error(f"❌ Erreur : {e}")
else:
    st.info("Veuillez charger le fichier principal pour commencer.")
