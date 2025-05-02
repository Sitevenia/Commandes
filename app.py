import streamlit as st
import pandas as pd
import numpy as np
import io
import logging
import re # Import regular expressions for sanitizing names

# --- (Keep existing functions: safe_read_excel, calculer_quantite_a_commander) ---

# Setup basic logging (optional)
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def safe_read_excel(uploaded_file, sheet_name, **kwargs):
    """Safely reads an Excel sheet, returning None if sheet not found."""
    try:
        # Ensure BytesIO is seekable if passed directly
        if isinstance(uploaded_file, io.BytesIO):
            uploaded_file.seek(0)
        return pd.read_excel(uploaded_file, sheet_name=sheet_name, **kwargs)
    except ValueError as e:
        # ValueError can be raised if sheet_name doesn't exist
        logging.warning(f"Sheet '{sheet_name}' not found or error reading it: {e}")
        st.warning(f"⚠️ L'onglet '{sheet_name}' n'a pas été trouvé dans le fichier Excel. Les vérifications associées seront ignorées.")
        return None
    except Exception as e:
        logging.error(f"Unexpected error reading sheet '{sheet_name}': {e}")
        st.error(f"❌ Erreur inattendue lors de la lecture de l'onglet '{sheet_name}'.")
        return None

def calculer_quantite_a_commander(df, semaine_columns, montant_minimum_input, duree_semaines):
    """
    Calcule la quantité à commander pour chaque produit.
    (Function remains the same as the previous version)
    """
    try:
        # --- Validation des Entrées ---
        if not isinstance(df, pd.DataFrame) or df.empty:
            st.error("Le DataFrame d'entrée est vide ou invalide pour le calcul.")
            return None
        required_cols = ["Stock", "Conditionnement", "Tarif d'achat"] + semaine_columns
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            st.error(f"Colonnes manquantes dans le DataFrame pour le calcul: {', '.join(missing_cols)}")
            return None
        if not semaine_columns:
            st.error("La liste des colonnes de semaines de vente est vide pour le calcul.")
            return None

        # Assurer que les colonnes nécessaires sont numériques et gérer les NaN/Infs
        df_calc = df.copy() # Work on a copy
        for col in required_cols:
            df_calc[col] = pd.to_numeric(df_calc[col], errors='coerce').replace([np.inf, -np.inf], np.nan).fillna(0)

        # --- Calculs des Ventes Moyennes ---
        num_semaines_totales = len(semaine_columns)
        ventes_N1 = df_calc[semaine_columns].sum(axis=1)

        # N-1 Calcs
        if num_semaines_totales >= 64:
            ventes_12_semaines_N1 = df_calc[semaine_columns[-64:-52]].sum(axis=1)
            ventes_12_semaines_N1_suivantes = df_calc[semaine_columns[-52:-40]].sum(axis=1)
            avg_12_N1 = ventes_12_semaines_N1 / 12
            avg_12_N1_suivantes = ventes_12_semaines_N1_suivantes / 12
        else:
            ventes_12_semaines_N1 = pd.Series(0, index=df_calc.index)
            ventes_12_semaines_N1_suivantes = pd.Series(0, index=df_calc.index)
            avg_12_N1 = 0
            avg_12_N1_suivantes = 0

        # Recent Calcs
        nb_semaines_recentes = min(num_semaines_totales, 12)
        if nb_semaines_recentes > 0:
            ventes_12_dernieres_semaines = df_calc[semaine_columns[-nb_semaines_recentes:]].sum(axis=1)
            avg_12_dernieres = ventes_12_dernieres_semaines / nb_semaines_recentes
        else:
            ventes_12_dernieres_semaines = pd.Series(0, index=df_calc.index)
            avg_12_dernieres = 0

        # --- Quantité Pondérée & Nécessaire ---
        quantite_ponderee = (0.5 * avg_12_dernieres + 0.2 * avg_12_N1 + 0.3 * avg_12_N1_suivantes)
        quantite_necessaire = quantite_ponderee * duree_semaines
        quantite_a_commander_series = (quantite_necessaire - df_calc["Stock"]).apply(lambda x: max(0, x))

        # --- Ajustements Basés sur les Règles ---
        conditionnement = df_calc["Conditionnement"]
        stock_actuel = df_calc["Stock"]
        tarif_achat = df_calc["Tarif d'achat"]
        quantite_a_commander = quantite_a_commander_series.tolist()

        # Cond
        for i in range(len(quantite_a_commander)):
            cond = conditionnement.iloc[i]
            q = quantite_a_commander[i]
            if q > 0 and cond > 0: quantite_a_commander[i] = int(np.ceil(q / cond) * cond)
            elif q > 0: quantite_a_commander[i] = 0
            else: quantite_a_commander[i] = 0

        # R1
        if nb_semaines_recentes > 0:
            for i in range(len(quantite_a_commander)):
                cond = conditionnement.iloc[i]
                ventes_recentes_count = (df_calc[semaine_columns[-nb_semaines_recentes:]].iloc[i] > 0).sum()
                if ventes_recentes_count >= 2 and stock_actuel.iloc[i] <= 1 and cond > 0:
                    quantite_a_commander[i] = max(quantite_a_commander[i], cond)

        # R2
        for i in range(len(quantite_a_commander)):
            ventes_tot_n1 = ventes_N1.iloc[i]; ventes_recentes_sum = ventes_12_dernieres_semaines.iloc[i]
            if ventes_tot_n1 < 6 and ventes_recentes_sum < 2: quantite_a_commander[i] = 0

        # --- Ajustement pour Montant Minimum Input ---
        montant_total_avant_ajust_min = sum(q * p for q, p in zip(quantite_a_commander, tarif_achat))
        if montant_minimum_input > 0 and montant_total_avant_ajust_min < montant_minimum_input:
            montant_total_actuel = montant_total_avant_ajust_min
            indices_commandes = [i for i, q in enumerate(quantite_a_commander) if q > 0]
            idx_pointer = 0; max_iterations = len(df_calc) * 10; iterations = 0
            while montant_total_actuel < montant_minimum_input and iterations < max_iterations:
                iterations += 1
                if not indices_commandes: break
                current_idx = indices_commandes[idx_pointer % len(indices_commandes)]
                cond = conditionnement.iloc[current_idx]; prix = tarif_achat.iloc[current_idx]
                if cond > 0 and prix > 0:
                    quantite_a_commander[current_idx] += cond; montant_total_actuel += cond * prix
                elif cond <= 0 :
                    indices_commandes.pop(idx_pointer % len(indices_commandes))
                    if not indices_commandes: continue
                    idx_pointer -= 1
                idx_pointer += 1
            if iterations >= max_iterations and montant_total_actuel < montant_minimum_input:
                 logging.error(f"Ajustement du montant minimum ({montant_minimum_input:.2f}€) échoué.")
                 st.error("L'ajustement automatique pour atteindre le montant minimum a échoué.")

        # --- Montant Final ---
        montant_total_final = sum(q * p for q, p in zip(quantite_a_commander, tarif_achat))
        return (quantite_a_commander, ventes_N1, ventes_12_semaines_N1, ventes_12_dernieres_semaines, montant_total_final)

    except KeyError as e: st.error(f"Erreur clé: '{e}'."); logging.error(f"KeyError calc: {e}"); return None
    except ValueError as e: st.error(f"Erreur valeur calc: {e}"); logging.error(f"ValueError calc: {e}"); return None
    except Exception as e: st.error(f"Erreur calc: {e}"); logging.exception("Error calc:"); return None


def sanitize_sheet_name(name):
    """Removes invalid characters for Excel sheet names and truncates."""
    if not isinstance(name, str):
        name = str(name)
    # Remove specific invalid characters: []:*?/\\
    sanitized = re.sub(r'[\[\]:*?/\\<>|"]', '_', name)
    # Sheet names cannot start or end with an apostrophe
    if sanitized.startswith("'"):
        sanitized = "_" + sanitized[1:]
    if sanitized.endswith("'"):
        sanitized = sanitized[:-1] + "_"
    # Truncate to maximum length (31 characters)
    return sanitized[:31]

# --- Streamlit App ---
st.set_page_config(page_title="Forecast App", layout="wide")
st.title("📦 Application de Prévision des Commandes")

uploaded_file = st.file_uploader("📁 Charger le fichier Excel principal", type=["xlsx", "xls"])

df_full = None
min_order_dict = {}

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

            if not semaine_columns: st.warning("⚠️ Aucune colonne de ventes hebdo identifiée.")

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
            result = calculer_quantite_a_commander(df_filtered, semaine_columns, montant_minimum_input_val, duree_semaines)

            if result is not None:
                st.success("✅ Calculs terminés.")
                (quantite_calculée, ventes_N1_calc, ventes_12_N1_calc,
                 ventes_12_last_calc, montant_total_calc) = result

                df_filtered.loc[:, "Quantité à commander"] = quantite_calculée
                df_filtered.loc[:, "Ventes N-1"] = ventes_N1_calc
                df_filtered.loc[:, "Ventes 12 semaines identiques N-1"] = ventes_12_N1_calc
                df_filtered.loc[:, "Ventes 12 dernières semaines"] = ventes_12_last_calc
                df_filtered.loc[:, "Total"] = df_filtered["Tarif d'achat"] * df_filtered["Quantité à commander"]
                df_filtered.loc[:, "Stock à terme"] = df_filtered["Stock"] + df_filtered["Quantité à commander"]

                st.metric(label="💰 Montant total GLOBAL calculé", value=f"{montant_total_calc:.2f} €")

                # --- MINIMUM WARNING (for single supplier selection only) ---
                if len(selected_fournisseurs) == 1:
                    selected_supplier = selected_fournisseurs[0]
                    if selected_supplier in min_order_dict:
                        required_minimum = min_order_dict[selected_supplier]
                        # Calculate the actual total for THIS supplier from the results
                        supplier_actual_total = df_filtered[df_filtered["Fournisseur"] == selected_supplier]["Total"].sum()
                        if required_minimum > 0 and supplier_actual_total < required_minimum:
                            diff = required_minimum - supplier_actual_total
                            st.warning(
                                f"⚠️ **Minimum Non Atteint (Fournisseur: {selected_supplier})**\n"
                                f"Montant Calculé: **{supplier_actual_total:.2f} €** | Minimum Requis: **{required_minimum:.2f} €** (Manque: {diff:.2f} €)\n\n"
                                f"➡️ Suggestion: Pour ajuster, modifiez le 'Montant minimum global (€)' à **{required_minimum:.2f}** et relancez."
                            )

                # --- Display Results Table (Combined) ---
                st.subheader("📊 Résultats Combinés")
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

                # --- EXPORT LOGIC (Modified for multi-sheet) ---
                st.subheader("⬇️ Exportation Excel par Fournisseur")
                # Filter *once* for items with quantity > 0 across *all* selected suppliers
                df_export_all = df_filtered[df_filtered["Quantité à commander"] > 0].copy()

                if not df_export_all.empty:
                    output = io.BytesIO()
                    try:
                        with pd.ExcelWriter(output, engine="openpyxl") as writer:
                            # Define columns for export sheets (can be same as display)
                            # Ensure 'Total' is the last numeric column for summary rows placement
                            export_columns = display_columns # Use the same columns as the combined display

                            for supplier in selected_fournisseurs:
                                # Filter for the current supplier
                                df_supplier_export = df_export_all[df_export_all["Fournisseur"] == supplier]

                                if not df_supplier_export.empty:
                                    # Prepare the main data part for the sheet
                                    df_supplier_sheet_data = df_supplier_export[export_columns].copy()

                                    # --- Create Summary Rows ---
                                    # 1. Calculate supplier total
                                    supplier_total = df_supplier_sheet_data["Total"].sum()

                                    # 2. Get supplier minimum from dict
                                    required_minimum = min_order_dict.get(supplier, 0) # Default to 0 if not found
                                    min_formatted = f"{required_minimum:.2f} €" if required_minimum > 0 else "N/A"

                                    # 3. Build Total Row DataFrame
                                    total_row_dict = {col: "" for col in export_columns}
                                    # Place labels in appropriate columns (e.g., Désignation)
                                    label_col = "Désignation Article" if "Désignation Article" in export_columns else export_columns[2] # Fallback column
                                    value_col = "Total" # Column for numeric total and formatted minimum

                                    total_row_dict[label_col] = "TOTAL COMMANDE"
                                    total_row_dict[value_col] = supplier_total # Store the actual number
                                    total_row_df = pd.DataFrame([total_row_dict])

                                    # 4. Build Minimum Row DataFrame
                                    min_row_dict = {col: "" for col in export_columns}
                                    min_row_dict[label_col] = "Minimum Requis"
                                    min_row_dict[value_col] = min_formatted # Store the formatted string
                                    min_row_df = pd.DataFrame([min_row_dict])

                                    # 5. Concatenate data + total row + minimum row
                                    df_sheet = pd.concat([df_supplier_sheet_data, total_row_df, min_row_df], ignore_index=True)

                                    # 6. Sanitize sheet name
                                    sanitized_name = sanitize_sheet_name(supplier)

                                    # 7. Write to Excel sheet
                                    df_sheet.to_excel(writer, sheet_name=sanitized_name, index=False)

                                    # --- Optional: Apply formatting using openpyxl ---
                                    # worksheet = writer.sheets[sanitized_name]
                                    # from openpyxl.styles import Font, Alignment, NumberFormat
                                    # # Example: Make summary rows bold
                                    # bold_font = Font(bold=True)
                                    # last_row_idx = len(df_sheet) # Index of Minimum row (1-based for openpyxl)
                                    # total_row_idx = last_row_idx - 1
                                    # value_col_letter = openpyxl.utils.get_column_letter(export_columns.index(value_col) + 1)
                                    #
                                    # for row_idx in [total_row_idx, last_row_idx]:
                                    #    for col_idx in range(1, len(export_columns) + 1):
                                    #        cell = worksheet.cell(row=row_idx, column=col_idx)
                                    #        cell.font = bold_font
                                    # # Format the actual total value as currency
                                    # total_cell = worksheet[f"{value_col_letter}{total_row_idx}"]
                                    # total_cell.number_format = '#,##0.00 €'

                                else:
                                     logging.info(f"No items to order for supplier '{supplier}', skipping sheet creation.")


                        output.seek(0) # Reset buffer position

                        # Create filename
                        suppliers_str = "multiples" if len(selected_fournisseurs) > 1 else sanitize_sheet_name(selected_fournisseurs[0])
                        filename = f"commande_{suppliers_str}_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx"

                        st.download_button(
                            label=f"📥 Télécharger Commandes ({len([s for s in selected_fournisseurs if not df_export_all[df_export_all['Fournisseur'] == s].empty])} Onglets)",
                            data=output,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    except Exception as e:
                        st.error(f"❌ Erreur lors de la création du fichier Excel : {e}")
                        logging.exception("Error during Excel export generation:")

                else:
                    st.info("ℹ️ Aucune quantité à commander pour l'exportation.")

            else:
                st.error("❌ Le calcul n'a pas abouti.")

        # --- Conditions for no calculation ---
        elif not selected_fournisseurs: st.warning("⚠️ Veuillez sélectionner au moins un fournisseur.")
        elif not semaine_columns and not df.empty: st.warning("⚠️ Calcul impossible: pas de colonnes ventes ou données filtrées incomplètes.")

    # --- File Loading Errors ---
    elif uploaded_file and df_full is None: st.error("❌ Échec lecture 'Tableau final'.")
    elif not uploaded_file: st.info("👋 Bienvenue ! Chargez votre fichier Excel.")
