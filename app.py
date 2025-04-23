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

        # Export Simulation 2
        output2 = io.BytesIO()
        with pd.ExcelWriter(output2, engine="xlsxwriter") as writer:
            # Filtrer les colonnes avant l'exportation
            df_filtered_sim2 = df_sim2[["Référence fournisseur", "Référence produit", "Désignation", "Qté Sim 2", "Montant Sim 2"] + selected_months]
            df_filtered_sim2.to_excel(writer, sheet_name="Simulation_2", index=False)
        output2.seek(0)
        st.download_button("📥 Télécharger Simulation 2", output2, file_name="simulation_2.xlsx")
