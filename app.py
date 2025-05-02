import streamlit as st
import pandas as pd
import numpy as np
import io
import logging
import re
import openpyxl # Ensure openpyxl is imported if we add formatting later

# --- (Keep existing functions: safe_read_excel, calculer_quantite_a_commander, sanitize_sheet_name) ---
# Make sure sanitize_sheet_name is defined as before:
def sanitize_sheet_name(name):
    """Removes invalid characters for Excel sheet names and truncates."""
    if not isinstance(name, str): name = str(name)
    sanitized = re.sub(r'[\[\]:*?/\\<>|"]', '_', name)
    if sanitized.startswith("'"): sanitized = "_" + sanitized[1:]
    if sanitized.endswith("'"): sanitized = sanitized[:-1] + "_"
    return sanitized[:31]

# --- Streamlit App ---
st.set_page_config(page_title="Forecast App", layout="wide")
st.title("📦 Application de Prévision des Commandes")

uploaded_file = st.file_uploader("📁 Charger le fichier Excel principal", type=["xlsx", "xls"])

# Initialize variables
df_full = None
min_order_dict = {}
# ...(rest of the initial file loading and minimum order dictionary creation code)...
# Assume df_full and min_order_dict are populated correctly here

if uploaded_file:
    file_buffer = io.BytesIO(uploaded_file.getvalue())
    logging.info("Attempting to read 'Tableau final' sheet.")
    df_full = safe_read_excel(file_buffer, sheet_name="Tableau final", header=7)

    logging.info("Attempting to read 'Minimum de commande' sheet.")
    df_min_commande = safe_read_excel(file_buffer, sheet_name="Minimum de commande")

    if df_min_commande is not None:
        logging.info("Processing 'Minimum de commande' sheet.")
        supplier_col_min = "Fournisseur"
        min_amount_col = "Minimum de Commande"
        required_min_cols = [supplier_col_min, min_amount_col]

        if all(col in df_min_commande.columns for col in required_min_cols):
            try:
                df_min_commande[supplier_col_min] = df_min_commande[supplier_col_min].astype(str).str.strip()
                df_min_commande[min_amount_col] = pd.to_numeric(df_min_commande[min_amount_col], errors='coerce')
                min_order_dict = df_min_commande.dropna(subset=[supplier_col_min, min_amount_col])\
                                               .set_index(supplier_col_min)[min_amount_col]\
                                               .to_dict()
                logging.info(f"Created minimum order dict: {len(min_order_dict)} entries.")
            except Exception as e:
                 st.error(f"❌ Erreur traitement 'Minimum de commande': {e}")
                 logging.exception("Error processing min order sheet:")
                 min_order_dict = {}
        else:
            missing_min_cols = [col for col in required_min_cols if col not in df_min_commande.columns]
            st.warning(f"⚠️ Colonnes manquantes ({', '.join(missing_min_cols)}) dans 'Minimum de commande'.")
            logging.warning(f"Missing columns in 'Minimum de commande': {missing_min_cols}")

    if df_full is not None:
        st.success("✅ Fichier principal ('Tableau final') chargé.")
        try:
            df = df_full[
                (df_full["Fournisseur"].notna()) & (df_full["Fournisseur"] != "") & (df_full["Fournisseur"] != "#FILTER") &
                (df_full["AF_RefFourniss"].notna()) & (df_full["AF_RefFourniss"] != "")
            ].copy()

            if df.empty:
                 st.warning("Aucune ligne valide après filtrage initial.")
                 fournisseurs = []
            else:
                fournisseurs = sorted(df["Fournisseur"].unique().tolist())
        except KeyError as e:
            st.error(f"❌ Colonne essentielle '{e}' manquante dans 'Tableau final'.")
            st.stop()

        selected_fournisseurs = st.multiselect("👤 Sélectionnez le(s) fournisseur(s)", options=fournisseurs, default=[])

        if selected_fournisseurs:
            df_filtered = df[df["Fournisseur"].isin(selected_fournisseurs)].copy()
        else:
            df_filtered = pd.DataFrame(columns=df.columns)

        # --- Identify Week Columns & Prepare ---
        start_col_index = 12
        semaine_columns = []
        if len(df_filtered.columns) > start_col_index:
            potential_week_cols = df_filtered.columns[start_col_index:].tolist()
            exclude_columns = ["Tarif d'achat", "Conditionnement", "Stock", "Total", "Stock à terme",
                               "Ventes N-1", "Ventes 12 semaines identiques N-1", "Ventes 12 dernières semaines",
                               "Quantité à commander", "Fournisseur"] # Also exclude Fournisseur here

            semaine_columns = [
                col for col in potential_week_cols
                if col not in exclude_columns and pd.api.types.is_numeric_dtype(df_filtered.get(col, pd.Series(dtype=float)).dtype)
            ]

            if not semaine_columns and not df_filtered.empty : st.warning("⚠️ Aucune colonne de ventes hebdo identifiée.")

            essential_numeric_cols = ["Stock", "Conditionnement", "Tarif d'achat"]
            for col in essential_numeric_cols:
                 if col in df_filtered.columns:
                     df_filtered[col] = pd.to_numeric(df_filtered[col], errors='coerce').fillna(0)
                 elif not df_filtered.empty:
                     st.error(f"Colonne essentielle '{col}' manquante."); st.stop()
        elif not df_filtered.empty:
            st.warning("Pas de colonnes après index 12 pour les ventes.")

        # --- Parameters ---
        col1, col2 = st.columns(2)
        with col1: duree_semaines = st.number_input("⏳ Durée couverture (semaines)", 4, 1, key="duree")
        with col2: montant_minimum_input_val = st.number_input("💶 Montant minimum global (€)", 0.0, 0.0, 50.0, "%.2f", key="montant_min")

        # --- Execute Calculation ---
        if not df_filtered.empty and semaine_columns:
            st.info("🚀 Lancement du calcul...")
            # Ensure calculation function is defined above or imported
            result = calculer_quantite_a_commander(df_filtered, semaine_columns, montant_minimum_input_val, duree_semaines)

            if result is not None:
                st.success("✅ Calculs terminés.")
                # --- (Code to unpack results and add columns to df_filtered) ---
                (quantite_calculée, ventes_N1_calc, ventes_12_N1_calc,
                 ventes_12_last_calc, montant_total_calc) = result

                df_filtered.loc[:, "Quantité à commander"] = quantite_calculée
                df_filtered.loc[:, "Ventes N-1"] = ventes_N1_calc
                df_filtered.loc[:, "Ventes 12 semaines identiques N-1"] = ventes_12_N1_calc
                df_filtered.loc[:, "Ventes 12 dernières semaines"] = ventes_12_last_calc
                df_filtered.loc[:, "Total"] = df_filtered["Tarif d'achat"] * df_filtered["Quantité à commander"]
                df_filtered.loc[:, "Stock à terme"] = df_filtered["Stock"] + df_filtered["Quantité à commander"]

                st.metric(label="💰 Montant total GLOBAL calculé", value=f"{montant_total_calc:.2f} €")

                # --- (Code for Minimum Warning - unchanged) ---
                if len(selected_fournisseurs) == 1:
                    # ... (warning logic as before) ...
                    selected_supplier = selected_fournisseurs[0]
                    if selected_supplier in min_order_dict:
                        required_minimum = min_order_dict[selected_supplier]
                        supplier_actual_total = df_filtered[df_filtered["Fournisseur"] == selected_supplier]["Total"].sum()
                        if required_minimum > 0 and supplier_actual_total < required_minimum:
                            diff = required_minimum - supplier_actual_total
                            st.warning(
                                f"⚠️ **Minimum Non Atteint (Fournisseur: {selected_supplier})**\n"
                                f"Montant Calculé: **{supplier_actual_total:.2f} €** | Minimum Requis: **{required_minimum:.2f} €** (Manque: {diff:.2f} €)\n\n"
                                f"➡️ Suggestion: Pour ajuster, modifiez le 'Montant minimum global (€)' à **{required_minimum:.2f}** et relancez."
                            )


                # --- (Code for Display Results Table - unchanged) ---
                st.subheader("📊 Résultats Combinés")
                # ... (dataframe display logic as before) ...
                required_display_columns = ["Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article", "Stock"]
                display_columns_base = required_display_columns + [
                    "Ventes N-1", "Ventes 12 semaines identiques N-1", "Ventes 12 dernières semaines",
                    "Conditionnement", "Quantité à commander", "Stock à terme",
                    "Tarif d'achat", "Total"
                ]
                display_columns = [col for col in display_columns_base if col in df_filtered.columns]
                missing_display_columns = [col for col in required_display_columns if col not in df_filtered.columns]

                if missing_display_columns:
                    st.error(f"❌ Colonnes manquantes affichage: {', '.join(missing_display_columns)}")
                else:
                    st.dataframe(df_filtered[display_columns].style.format({
                        "Tarif d'achat": "{:.2f}€", "Total": "{:.2f}€",
                        "Ventes N-1": "{:,.0f}", "Ventes 12 semaines identiques N-1": "{:,.0f}",
                        "Ventes 12 dernières semaines": "{:,.0f}", "Stock": "{:,.0f}",
                        "Conditionnement": "{:,.0f}", "Quantité à commander": "{:,.0f}",
                        "Stock à terme": "{:,.0f}"
                    }, na_rep="-"))


                # --- EXPORT LOGIC (Refined Multi-Sheet) ---
                st.subheader("⬇️ Exportation Excel par Fournisseur")
                df_export_all = df_filtered[df_filtered["Quantité à commander"] > 0].copy()

                if not df_export_all.empty:
                    output = io.BytesIO()
                    sheets_created_count = 0 # Counter for actual sheets written
                    try:
                        # Use ExcelWriter context manager
                        with pd.ExcelWriter(output, engine="openpyxl") as writer:
                            logging.info(f"Starting export loop for suppliers: {selected_fournisseurs}")
                            # Define columns for export sheets - IMPORTANT: Ensure 'Fournisseur' is NOT included here if you want it removed from individual sheets
                            export_columns = [col for col in display_columns if col != 'Fournisseur'] # Remove Fournisseur column from individual sheets

                            for supplier in selected_fournisseurs:
                                logging.info(f"Processing supplier: {supplier}")
                                # Filter for the current supplier's items with quantity > 0
                                df_supplier_export = df_export_all[df_export_all["Fournisseur"] == supplier]

                                if not df_supplier_export.empty:
                                    logging.info(f"Found {len(df_supplier_export)} items to order for {supplier}.")
                                    # Prepare the main data part for the sheet, selecting only the desired export columns
                                    df_supplier_sheet_data = df_supplier_export[export_columns].copy()

                                    # --- Create Summary Rows ---
                                    supplier_total = df_supplier_sheet_data["Total"].sum()
                                    required_minimum = min_order_dict.get(supplier, 0)
                                    min_formatted = f"{required_minimum:.2f} €" if required_minimum > 0 else "N/A"

                                    # Determine columns for labels and values in summary rows
                                    label_col = "Désignation Article" if "Désignation Article" in export_columns else export_columns[1] # Use 2nd col as fallback
                                    value_col = "Total" if "Total" in export_columns else export_columns[-1] # Use last col as fallback

                                    # Build Total Row DataFrame
                                    total_row_dict = {col: "" for col in export_columns}
                                    total_row_dict[label_col] = "TOTAL COMMANDE"
                                    total_row_dict[value_col] = supplier_total # Numeric value
                                    total_row_df = pd.DataFrame([total_row_dict])

                                    # Build Minimum Row DataFrame
                                    min_row_dict = {col: "" for col in export_columns}
                                    min_row_dict[label_col] = "Minimum Requis"
                                    min_row_dict[value_col] = min_formatted # Formatted string
                                    min_row_df = pd.DataFrame([min_row_dict])

                                    # Concatenate data + total row + minimum row
                                    df_sheet = pd.concat([df_supplier_sheet_data, total_row_df, min_row_df], ignore_index=True)

                                    # Sanitize sheet name
                                    sanitized_name = sanitize_sheet_name(supplier)
                                    logging.info(f"Using sanitized sheet name: {sanitized_name}")

                                    # Write to Excel sheet within the context manager
                                    try:
                                        df_sheet.to_excel(writer, sheet_name=sanitized_name, index=False)
                                        sheets_created_count += 1
                                        logging.info(f"Successfully wrote sheet: {sanitized_name}")
                                    except Exception as write_error:
                                        st.error(f"❌ Erreur lors de l'écriture de l'onglet pour {supplier} ({sanitized_name}): {write_error}")
                                        logging.error(f"Error writing sheet {sanitized_name} for supplier {supplier}: {write_error}")

                                else:
                                     logging.info(f"No items to order for supplier '{supplier}', skipping sheet creation.")
                                     # st.caption(f"Aucun article à commander pour {supplier}, onglet non créé.") # Optional user feedback

                        # After the 'with' block, the Excel file is saved to the buffer 'output'
                        output.seek(0) # Reset buffer position for reading

                        if sheets_created_count > 0:
                            # Create filename
                            suppliers_str = "multiples" if len(selected_fournisseurs) > 1 else sanitize_sheet_name(selected_fournisseurs[0])
                            filename = f"commande_{suppliers_str}_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx"

                            st.download_button(
                                label=f"📥 Télécharger Commandes ({sheets_created_count} Onglet{'s' if sheets_created_count > 1 else ''})", # Dynamic label
                                data=output,
                                file_name=filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                            logging.info(f"Download button created for {sheets_created_count} sheets.")
                        else:
                            st.info("ℹ️ Aucune quantité à commander trouvée pour l'exportation pour les fournisseurs sélectionnés.")
                            logging.info("No sheets were created as no items had quantity > 0.")

                    # Catch errors during the ExcelWriter process itself
                    except Exception as e:
                        st.error(f"❌ Erreur majeure lors de la création du fichier Excel : {e}")
                        logging.exception("Error during ExcelWriter context or processing:")

                else:
                    st.info("ℹ️ Aucune quantité à commander globale trouvée pour l'exportation.")
                    logging.info("df_export_all was empty, skipping export.")

            # --- (Error handling for calculation failure) ---
            else:
                st.error("❌ Le calcul n'a pas abouti.")

        # --- (Conditions for no calculation / no selection / no columns) ---
        elif not selected_fournisseurs: st.warning("⚠️ Veuillez sélectionner au moins un fournisseur.")
        elif not semaine_columns and not df_filtered.empty: st.warning("⚠️ Calcul impossible: pas de colonnes ventes ou données filtrées incomplètes.")

    # --- (File Loading Errors) ---
    elif uploaded_file and df_full is None: st.error("❌ Échec lecture 'Tableau final'.")
    elif not uploaded_file: st.info("👋 Bienvenue ! Chargez votre fichier Excel.")
