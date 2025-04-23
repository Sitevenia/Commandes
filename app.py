import streamlit as st
import pandas as pd
import numpy as np
import io

def repartir_et_ajuster(total_qte, ventes_n1_mois, conditionnement):
    """Répartit une quantité totale selon la saisonnalité et ajuste aux conditionnements."""
    if total_qte <= 0 or conditionnement is None or conditionnement <= 0:
        return [0] * len(ventes_n1_mois)

    total_ventes = sum(ventes_n1_mois)
    if total_ventes == 0:
        repartition = [1] * len(ventes_n1_mois)
        total_ventes = len(ventes_n1_mois)
    else:
        repartition = ventes_n1_mois

    proportions = [v / total_ventes for v in repartition]
    qtes_mensuelles = [round(total_qte * p) for p in proportions]

    # Ajuster chaque mois au multiple de conditionnement
    qtes_conditionnees = [int(round(q / conditionnement)) * conditionnement for q in qtes_mensuelles]

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

def repartir_et_ajuster(qte_totale, saisonnalite, conditionnement):
    try:
        if not np.isfinite(qte_totale) or qte_totale <= 0 or saisonnalite.isnull().all():
            return [0] * len(saisonnalite)
        saisonnalite = saisonnalite.fillna(0)
        if saisonnalite.sum() == 0:
            return [0] * len(saisonnalite)
        saisonnalite = saisonnalite / saisonnalite.sum()
        raw = qte_totale * saisonnalite
        repartition = np.ceil(raw / conditionnement) * conditionnement
        repartition = np.nan_to_num(repartition, nan=0.0, posinf=0.0, neginf=0.0)
        ecart = int(qte_totale - repartition.sum())
        while ecart != 0:
            idx = np.argmax(saisonnalite) if ecart > 0 else np.argmax(repartition)
            modif = conditionnement if ecart > 0 else -conditionnement
            tentative = repartition[idx] + modif
            if tentative >= 0:
                repartition[idx] = tentative
                ecart -= modif
            else:
                break
        return [safe_int(x) for x in repartition]
    except:
        return [0] * len(saisonnalite)

# Chargement du fichier principal
uploaded_file = st.file_uploader("📁 Charger le fichier Excel principal", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, sheet_name="Tableau final")
        st.success("✅ Fichier principal chargé avec succès.")

        month_columns = [str(i) for i in range(1, 13)]
        selected_months = st.multiselect("Sélectionnez les mois à inclure", month_columns, default=month_columns)

        for col in month_columns + ["Tarif d'achat", "Conditionnement"]:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

        df["Conditionnement"] = df["Conditionnement"].replace(0, 1)
        df["Total ventes N-1"] = df[selected_months].sum(axis=1).replace(0, np.nan)
        saisonnalite = df[selected_months].div(df["Total ventes N-1"], axis=0).replace([np.inf, -np.inf], 0).fillna(0)

        # Simulation 1
        st.subheader("Simulation 1 : progression personnalisée")
        progression = st.number_input("📈 Progression (%)", value=0.0, step=1.0)
        df["Qté Sim 1"] = df["Total ventes N-1"] * (1 + progression / 100)
        df["Qté Sim 1"] = (np.ceil(df["Qté Sim 1"] / df["Conditionnement"]) * df["Conditionnement"]).fillna(0).astype(int)

        if st.button("▶️ Lancer la Simulation 1"):
            for i in df.index:
                repartition = repartir_et_ajuster(
                    df.at[i, "Qté Sim 1"],
                    saisonnalite.loc[i, selected_months],
                    df.at[i, "Conditionnement"]
                )
                # Assurez-vous que la longueur de repartition correspond à celle des colonnes sélectionnées
                if len(repartition) == len(selected_months):
                    df.loc[i, selected_months] = repartition
                else:
                    st.error("Erreur : La longueur de la répartition ne correspond pas aux mois sélectionnés.")

            df["Montant Sim 1"] = df["Qté Sim 1"] * df["Tarif d'achat"]
            total_sim1 = df["Montant Sim 1"].sum()
            st.metric("💰 Total Simulation 1", f"€ {total_sim1:,.2f}")

            # Export Simulation 1
            output1 = io.BytesIO()
            with pd.ExcelWriter(output1, engine="xlsxwriter") as writer:
                # Filtrer les colonnes avant l'exportation
                df_filtered = df[["Référence fournisseur", "Référence produit", "Désignation", "Qté Sim 1", "Montant Sim 1"] + selected_months]
                df_filtered.to_excel(writer, sheet_name="Simulation_1", index=False)
            output1.seek(0)
            st.download_button("📥 Télécharger Simulation 1", output1, file_name="simulation_1.xlsx")

        # Simulation 2
        st.subheader("Simulation 2 : objectif d'achat ajusté précisément")
        objectif = st.number_input("🎯 Objectif (€)", value=0.0, step=1000.0)

        if objectif > 0:
            if st.button("▶️ Lancer la Simulation 2"):
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
                        saisonnalite.loc[i, selected_months],
                        df_sim2.at[i, "Conditionnement"]
                    )
                    # Assurez-vous que la longueur de repartition correspond à celle des colonnes sélectionnées
                    if len(repartition) == len(selected_months):
                        df_sim2.loc[i, selected_months] = repartition
                    else:
                        st.error("Erreur : La longueur de la répartition ne correspond pas aux mois sélectionnés.")

                df_sim2["Montant Sim 2"] = df_sim2["Qté Sim 2"] * df_sim2["Tarif d'achat"]
                total_sim2 = df_sim2["Montant Sim 2"].sum()
                st.metric("✅ Montant Simulation 2", f"€ {total_sim2:,.2f}")

                st.dataframe(df_sim2[["Référence fournisseur", "Référence produit", "Désignation", "Qté Sim 2", "Montant Sim 2"]])

                # Comparatif
                st.subheader("📊 Comparatif")
                comparatif = df[["Référence fournisseur", "Référence produit", "Désignation"]].copy()
                comparatif["Qté Sim 1"] = df["Qté Sim 1"]
                comparatif["Qté Sim 2"] = df_sim2["Qté Sim 2"]
                comparatif["Montant Sim 2"] = df_sim2["Montant Sim 2"]
                st.dataframe(comparatif)

                # Export Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                    # Filtrer les colonnes avant l'exportation
                    df_filtered_sim1 = df[["Référence fournisseur", "Référence produit", "Désignation", "Qté Sim 1"] + selected_months]
                    df_filtered_sim1.to_excel(writer, sheet_name="Simulation_1", index=False)

                    df_filtered_sim2 = df_sim2[["Référence fournisseur", "Référence produit", "Désignation", "Qté Sim 2", "Montant Sim 2"] + selected_months]
                    df_filtered_sim2.to_excel(writer, sheet_name="Simulation_2", index=False)

                    comparatif_filtered = comparatif[["Référence fournisseur", "Référence produit", "Désignation", "Qté Sim 1", "Qté Sim 2", "Montant Sim 2"]]
                    comparatif_filtered.to_excel(writer, sheet_name="Comparatif", index=False)
                output.seek(0)
                st.download_button("📥 Télécharger le fichier Excel", output, file_name="forecast_result_final.xlsx")

    except Exception as e:
        st.error(f"❌ Erreur : {e}")
else:
    st.info("Veuillez charger le fichier principal pour commencer.")
