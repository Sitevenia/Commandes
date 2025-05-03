import streamlit as st
import pandas as pd
import numpy as np
import io
import logging
import re
import openpyxl # Import openpyxl for direct manipulation
from openpyxl.utils import get_column_letter # Utility to get column letters

# --- Logging Configuration ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- Helper Functions ---
# (safe_read_excel, calculer_quantite_a_commander, sanitize_sheet_name remain the same)
def safe_read_excel(uploaded_file, sheet_name, **kwargs):
    """
    Safely reads an Excel sheet, returning None if sheet not found or error occurs.
    Handles BytesIO seeking.
    """
    try:
        if isinstance(uploaded_file, io.BytesIO):
            uploaded_file.seek(0)
        return pd.read_excel(uploaded_file, sheet_name=sheet_name, **kwargs)
    except ValueError as e:
        if f"Worksheet named '{sheet_name}' not found" in str(e):
             logging.warning(f"Sheet '{sheet_name}' not found.")
             st.warning(f"⚠️ L'onglet '{sheet_name}' n'a pas été trouvé.")
        else:
             logging.error(f"ValueError reading sheet '{sheet_name}': {e}")
             st.error(f"❌ Erreur de valeur lors de la lecture de l'onglet '{sheet_name}': {e}.")
        return None
    except FileNotFoundError:
        logging.error(f"FileNotFoundError reading sheet '{sheet_name}'.")
        st.error(f"❌ Fichier non trouvé (erreur interne) lors de la lecture '{sheet_name}'.")
        return None
    except Exception as e:
        logging.error(f"Unexpected error reading sheet '{sheet_name}': {type(e).__name__} - {e}")
        st.error(f"❌ Erreur inattendue ({type(e).__name__}) lecture '{sheet_name}': {e}.")
        return None

def calculer_quantite_a_commander(df, semaine_columns, montant_minimum_input, duree_semaines):
    """ Calcule la quantité à commander. (Code identical to previous version) """
    try:
        # --- Validation des Entrées ---
        if not isinstance(df, pd.DataFrame) or df.empty: return None
        required_cols = ["Stock", "Conditionnement", "Tarif d'achat"] + semaine_columns
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols: st.error(f"Colonnes manquantes: {', '.join(missing_cols)}"); return None
        if not semaine_columns: st.error("Colonnes semaines vides."); return None
        df_calc = df.copy()
        for col in required_cols: df_calc[col] = pd.to_numeric(df_calc[col], errors='coerce').replace([np.inf, -np.inf], np.nan).fillna(0)
        # --- Calculs Ventes Moyennes ---
        num_semaines_totales = len(semaine_columns)
        ventes_N1 = df_calc[semaine_columns].sum(axis=1)
        if num_semaines_totales >= 64:
            ventes_12_semaines_N1 = df_calc[semaine_columns[-64:-52]].sum(axis=1)
            ventes_12_semaines_N1_suivantes = df_calc[semaine_columns[-52:-40]].sum(axis=1)
            avg_12_N1 = ventes_12_semaines_N1 / 12; avg_12_N1_suivantes = ventes_12_semaines_N1_suivantes / 12
        else:
            ventes_12_semaines_N1 = pd.Series(0, index=df_calc.index); ventes_12_semaines_N1_suivantes = pd.Series(0, index=df_calc.index)
            avg_12_N1 = 0; avg_12_N1_suivantes = 0
        nb_semaines_recentes = min(num_semaines_totales, 12)
        if nb_semaines_recentes > 0:
            ventes_12_dernieres_semaines = df_calc[semaine_columns[-nb_semaines_recentes:]].sum(axis=1)
            avg_12_dernieres = ventes_12_dernieres_semaines / nb_semaines_recentes
        else:
            ventes_12_dernieres_semaines = pd.Series(0, index=df_calc.index); avg_12_dernieres = 0
        # --- Quantité Pondérée & Nécessaire ---
        quantite_ponderee = (0.5 * avg_12_dernieres + 0.2 * avg_12_N1 + 0.3 * avg_12_N1_suivantes)
        quantite_necessaire = quantite_ponderee * duree_semaines
        quantite_a_commander_series = (quantite_necessaire - df_calc["Stock"]).apply(lambda x: max(0, x))
        # --- Ajustements Règles ---
        conditionnement = df_calc["Conditionnement"]; stock_actuel = df_calc["Stock"]; tarif_achat = df_calc["Tarif d'achat"]
        quantite_a_commander = quantite_a_commander_series.tolist()
        for i in range(len(quantite_a_commander)): # Cond
            cond = conditionnement.iloc[i]; q = quantite_a_commander[i]
            if q > 0 and cond > 0: quantite_a_commander[i] = int(np.ceil(q / cond) * cond)
            elif q > 0: quantite_a_commander[i] = 0
            else: quantite_a_commander[i] = 0
        if nb_semaines_recentes > 0: # R1
            for i in range(len(quantite_a_commander)):
                cond = conditionnement.iloc[i]; ventes_recentes_count = (df_calc[semaine_columns[-nb_semaines_recentes:]].iloc[i] > 0).sum()
                if ventes_recentes_count >= 2 and stock_actuel.iloc[i] <= 1 and cond > 0: quantite_a_commander[i] = max(quantite_a_commander[i], cond)
        for i in range(len(quantite_a_commander)): # R2
            ventes_tot_n1 = ventes_N1.iloc[i]; ventes_recentes_sum = ventes_12_dernieres_semaines.iloc[i]
            if ventes_tot_n1 < 6 and ventes_recentes_sum < 2: quantite_a_commander[i] = 0
        # --- Ajustement Montant Min ---
        montant_total_avant_ajust_min = sum(q * p for q, p in zip(quantite_a_commander, tarif_achat))
        if montant_minimum_input > 0 and montant_total_avant_ajust_min < montant_minimum_input:
            montant_total_actuel = montant_total_avant_ajust_min
            indices_commandes = [i for i, q in enumerate(quantite_a_commander) if q > 0]
            idx_pointer = 0; max_iterations = len(df_calc) * 10; iterations = 0
            while montant_total_actuel < montant_minimum_input and iterations < max_iterations:
                iterations += 1;
                if not indices_commandes: break
                current_idx = indices_commandes[idx_pointer % len(indices_commandes)]; cond = conditionnement.iloc[current_idx]; prix = tarif_achat.iloc[current_idx]
                if cond > 0 and prix > 0: quantite_a_commander[current_idx] += cond; montant_total_actuel += cond * prix
                elif cond <= 0 : indices_commandes.pop(idx_pointer % len(indices_commandes));
                if not indices_commandes: continue; idx_pointer -= 1
                idx_pointer += 1
            if iterations >= max_iterations and montant_total_actuel < montant_minimum_input: st.error("Ajustement montant min échoué (max iter).")
        # --- Montant Final ---
        montant_total_final = sum(q * p for q, p in zip(quantite_a_commander, tarif_achat))
        return (quantite_a_commander, ventes_N1, ventes_12_semaines_N1, ventes_12_dernieres_semaines, montant_total_final)
    except KeyError as e: st.error(f"Erreur clé calc: '{e}'."); return None
    except ValueError as e: st.error(f"Erreur valeur calc: {e}"); return None
    except Exception as e: st.error(f"Erreur calc: {e}"); logging.exception("Error calc:"); return None

def sanitize_sheet_name(name):
    """Removes invalid characters for Excel sheet names and truncates."""
    if not isinstance(name, str): name = str(name)
    sanitized = re.sub(r'[\[\]:*?/\\<>|"]', '_', name)
    if sanitized.startswith("'"): sanitized = "_" + sanitized[1:]
    if sanitized.endswith("'"): sanitized = sanitized[:-1] + "_"
    return sanitized[:31]

# --- Streamlit App Main Logic ---
st.set_page_config(page_title="Forecast App", layout="wide")
st.title("📦 Application de Prévision des Commandes")

uploaded_file = st.file_uploader("📁 Charger le fichier Excel principal", type=["xlsx", "xls"])

# Initialize variables
df_full = None
min_order_dict = {}
df_min_commande = None

if uploaded_file:
    try: # Outer try block for initial loading
        file_buffer = io.BytesIO(uploaded_file.getvalue())
        st.info("Lecture onglet 'Tableau final'...")
        df_full = safe_read_excel(file_buffer, sheet_name="Tableau final", header=7)

        if df_full is None:
             st.error("❌ Échec lecture 'Tableau final'. Impossible de continuer.")
             st.stop()
        else:
             st.success("✅ Onglet 'Tableau final' lu.")

        # --- Read Minimum Order Sheet ---
        st.info("Lecture onglet 'Minimum de commande'...")
        df_min_commande = safe_read_excel(file_buffer, sheet_name="Minimum de commande")
        if df_min_commande is not None:
            st.success("✅ Onglet 'Minimum de commande' lu.")
            # Process Minimum Order Data
            supplier_col_min = "Fournisseur"; min_amount_col = "Minimum de Commande" # Adjust if needed
            required_min_cols = [supplier_col_min, min_amount_col]
            if all(col in df_min_commande.columns for col in required_min_cols):
                try:
                    df_min_commande[supplier_col_min] = df_min_commande[supplier_col_min].astype(str).str.strip()
                    df_min_commande[min_amount_col] = pd.to_numeric(df_min_commande[min_amount_col], errors='coerce')
                    min_order_dict = df_min_commande.dropna(subset=[supplier_col_min, min_amount_col])\
                                                .set_index(supplier_col_min)[min_amount_col].to_dict()
                    logging.info(f"Min order dict: {len(min_order_dict)} entries.")
                except Exception as e_min_proc: st.error(f"❌ Erreur traitement 'Min commande': {e_min_proc}")
            else: st.warning(f"⚠️ Colonnes manquantes ({', '.join(required_min_cols)}) dans 'Min commande'.")
        # else: safe_read_excel already displayed warning

    except Exception as e_load:
        st.error(f"❌ Erreur lecture fichier : {e_load}"); logging.exception("File loading error:"); st.stop()

    # --- Continue processing if df_full exists ---
    if df_full is not None:
        try: # Filter initial data
            df = df_full[
                (df_full["Fournisseur"].notna()) & (df_full["Fournisseur"] != "") & (df_full["Fournisseur"] != "#FILTER") &
                (df_full["AF_RefFourniss"].notna()) & (df_full["AF_RefFourniss"] != "")
            ].copy()
            fournisseurs = sorted(df["Fournisseur"].unique().tolist()) if not df.empty else []
            if df.empty and not df_full.empty: st.warning("Aucune ligne valide après filtrage initial.")
        except KeyError as e_filter: st.error(f"❌ Colonne filtrage '{e_filter}' manquante."); st.stop()

        # --- User Selections ---
        selected_fournisseurs = st.multiselect("👤 Sélectionnez fournisseur(s)", options=fournisseurs, default=[])
        df_filtered = df[df["Fournisseur"].isin(selected_fournisseurs)].copy() if selected_fournisseurs else pd.DataFrame(columns=df.columns)

        # --- Identify Week Columns & Prepare Data ---
        start_col_index = 12; semaine_columns = []
        if len(df_filtered.columns) > start_col_index:
            potential_week_cols = df_filtered.columns[start_col_index:].tolist()
            exclude_columns = ["Tarif d'achat", "Conditionnement", "Stock", "Total", "Stock à terme", "Ventes N-1", "Ventes 12 semaines identiques N-1", "Ventes 12 dernières semaines", "Quantité à commander", "Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article"]
            semaine_columns = [col for col in potential_week_cols if col not in exclude_columns and pd.api.types.is_numeric_dtype(df_filtered.get(col, pd.Series(dtype=float)).dtype)]
            if not semaine_columns and not df_filtered.empty: st.warning("⚠️ Aucune colonne vente hebdo trouvée.")
            essential_numeric_cols = ["Stock", "Conditionnement", "Tarif d'achat"]
            for col in essential_numeric_cols:
                 if col in df_filtered.columns: df_filtered[col] = pd.to_numeric(df_filtered[col], errors='coerce').fillna(0)
                 elif not df_filtered.empty: st.error(f"Colonne essentielle '{col}' manquante."); st.stop()
        elif not df_filtered.empty: st.warning("Pas de colonnes après index 12.")

        # --- Calculation Parameters ---
        col1, col2 = st.columns(2)
        with col1: duree_semaines = st.number_input("⏳ Durée couverture (semaines)", 4, 1, key="duree")
        with col2: montant_minimum_input_val = st.number_input("💶 Montant minimum global (€)", 0.0, 0.0, 50.0, "%.2f", key="montant_min")

        # --- Execute Calculation ---
        if not df_filtered.empty and semaine_columns:
            st.info("🚀 Lancement calcul...")
            result = calculer_quantite_a_commander(df_filtered, semaine_columns, montant_minimum_input_val, duree_semaines)

            if result is not None:
                st.success("✅ Calculs terminés.")
                (quantite_calculée, ventes_N1_calc, ventes_12_N1_calc, ventes_12_last_calc, montant_total_calc) = result
                df_filtered.loc[:, "Quantité à commander"] = quantite_calculée
                df_filtered.loc[:, "Ventes N-1"] = ventes_N1_calc
                df_filtered.loc[:, "Ventes 12 semaines identiques N-1"] = ventes_12_N1_calc
                df_filtered.loc[:, "Ventes 12 dernières semaines"] = ventes_12_last_calc
                df_filtered.loc[:, "Tarif d'achat"] = pd.to_numeric(df_filtered["Tarif d'achat"], errors='coerce').fillna(0) # Ensure Tariff is numeric before calc
                df_filtered.loc[:, "Total"] = df_filtered["Tarif d'achat"] * df_filtered["Quantité à commander"] # Recalculate Total based on final Qty
                df_filtered.loc[:, "Stock à terme"] = df_filtered["Stock"] + df_filtered["Quantité à commander"]

                st.metric(label="💰 Montant total GLOBAL calculé", value=f"{montant_total_calc:.2f} €")

                # --- MINIMUM WARNING (Single Supplier) ---
                if len(selected_fournisseurs) == 1:
                    selected_supplier = selected_fournisseurs[0]
                    if selected_supplier in min_order_dict:
                        required_minimum = min_order_dict[selected_supplier]
                        supplier_actual_total = df_filtered[df_filtered["Fournisseur"] == selected_supplier]["Total"].sum()
                        if required_minimum > 0 and supplier_actual_total < required_minimum:
                            diff = required_minimum - supplier_actual_total
                            st.warning(f"⚠️ **Minimum Non Atteint ({selected_supplier})**\nMontant Calculé: **{supplier_actual_total:,.2f} €** | Requis: **{required_minimum:,.2f} €** (Manque: {diff:,.2f} €)\n➡️ Suggestion: Modifiez 'Montant min global (€)' à **{required_minimum:.2f}** et relancez.")

                # --- Display Combined Results ---
                st.subheader("📊 Résultats Combinés")
                required_display_columns = ["Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article", "Stock"]
                display_columns_base = required_display_columns + ["Ventes N-1", "Ventes 12 semaines identiques N-1", "Ventes 12 dernières semaines", "Conditionnement", "Quantité à commander", "Stock à terme", "Tarif d'achat", "Total"]
                display_columns = [col for col in display_columns_base if col in df_filtered.columns]
                if any(col not in df_filtered.columns for col in required_display_columns): st.error("❌ Colonnes manquantes pour affichage.")
                else: st.dataframe(df_filtered[display_columns].style.format({"Tarif d'achat": "{:,.2f}€", "Total": "{:,.2f}€", "Ventes N-1": "{:,.0f}", "Ventes 12 semaines identiques N-1": "{:,.0f}", "Ventes 12 dernières semaines": "{:,.0f}", "Stock": "{:,.0f}", "Conditionnement": "{:,.0f}", "Quantité à commander": "{:,.0f}", "Stock à terme": "{:,.0f}"}, na_rep="-", thousands=","))

                # --- EXPORT LOGIC (Multi-Sheet with Formulas) ---
                st.subheader("⬇️ Exportation Excel par Fournisseur (avec formules)")
                df_export_all = df_filtered[df_filtered["Quantité à commander"] > 0].copy()

                if not df_export_all.empty:
                    output = io.BytesIO()
                    sheets_created_count = 0
                    try:
                        with pd.ExcelWriter(output, engine="openpyxl") as writer:
                            logging.info(f"Export: Processing suppliers: {selected_fournisseurs}")
                            # Define export columns (excluding 'Fournisseur' for individual sheets)
                            # IMPORTANT: Adjust col names if they differ in your df_filtered!
                            qty_col_name = "Quantité à commander"
                            price_col_name = "Tarif d'achat"
                            total_col_name = "Total"
                            export_columns = [col for col in display_columns if col != 'Fournisseur' and col in df_export_all.columns] # Ensure columns exist

                            # Verify essential columns for formula exist in export_columns
                            if not all(c in export_columns for c in [qty_col_name, price_col_name, total_col_name]):
                                st.error(f"❌ Colonnes essentielles ('{qty_col_name}', '{price_col_name}', '{total_col_name}') manquantes pour l'export avec formules.")
                                logging.error("Essential columns missing for formula export.")
                            else:
                                # Get 0-based indices of columns in the final export list
                                try:
                                    qty_col_idx = export_columns.index(qty_col_name)
                                    price_col_idx = export_columns.index(price_col_name)
                                    total_col_idx = export_columns.index(total_col_name)
                                    # Convert to Excel 1-based column letters
                                    qty_col_letter = get_column_letter(qty_col_idx + 1)
                                    price_col_letter = get_column_letter(price_col_idx + 1)
                                    total_col_letter = get_column_letter(total_col_idx + 1)
                                    formula_ready = True
                                    logging.info(f"Formula columns: Qty={qty_col_letter}, Price={price_col_letter}, Total={total_col_letter}")
                                except ValueError:
                                    st.error("Erreur interne: Impossible de trouver les indices des colonnes pour les formules.")
                                    logging.error("Could not find column indices for formulas.")
                                    formula_ready = False

                                if formula_ready:
                                    for supplier in selected_fournisseurs:
                                        logging.info(f"Export: Processing supplier {supplier}")
                                        df_supplier_export = df_export_all[df_export_all["Fournisseur"] == supplier]

                                        if not df_supplier_export.empty:
                                            df_supplier_sheet_data = df_supplier_export[export_columns].copy()
                                            num_data_rows = len(df_supplier_sheet_data)

                                            # --- Summary Rows (prep before writing) ---
                                            supplier_total_val = df_supplier_sheet_data[total_col_name].sum() # Calculated value
                                            required_minimum = min_order_dict.get(supplier, 0)
                                            min_formatted = f"{required_minimum:,.2f} €" if required_minimum > 0 else "N/A"
                                            label_col = "Désignation Article" if "Désignation Article" in export_columns else export_columns[1]
                                            value_col = total_col_name # Total column holds values/formulas

                                            total_row_dict = {col: "" for col in export_columns}; total_row_dict[label_col] = "TOTAL COMMANDE"; total_row_dict[value_col] = supplier_total_val # Placeholder value initially
                                            min_row_dict = {col: "" for col in export_columns}; min_row_dict[label_col] = "Minimum Requis"; min_row_dict[value_col] = min_formatted
                                            total_row_df = pd.DataFrame([total_row_dict]); min_row_df = pd.DataFrame([min_row_dict])
                                            df_sheet = pd.concat([df_supplier_sheet_data, total_row_df, min_row_df], ignore_index=True)

                                            sanitized_name = sanitize_sheet_name(supplier)
                                            try:
                                                # Write data (including placeholder total value)
                                                df_sheet.to_excel(writer, sheet_name=sanitized_name, index=False)

                                                # --- Get worksheet and apply FORMULAS ---
                                                worksheet = writer.sheets[sanitized_name]

                                                # Apply formula to data rows (Excel rows 2 to num_data_rows + 1)
                                                for row_num in range(2, num_data_rows + 2):
                                                    formula = f"={qty_col_letter}{row_num}*{price_col_letter}{row_num}"
                                                    worksheet[f"{total_col_letter}{row_num}"] = formula
                                                    # Optional: Apply currency format to formula cells
                                                    worksheet[f"{total_col_letter}{row_num}"].number_format = '#,##0.00 €'

                                                # Apply SUM formula to the "TOTAL COMMANDE" row
                                                total_formula_row_num = num_data_rows + 2 # Row number for TOTAL COMMANDE
                                                if num_data_rows > 0: # Only add sum if there are data rows
                                                     sum_formula = f"=SUM({total_col_letter}2:{total_col_letter}{num_data_rows + 1})"
                                                     worksheet[f"{total_col_letter}{total_formula_row_num}"] = sum_formula
                                                     # Optional: Apply currency format and bold font to SUM cell
                                                     worksheet[f"{total_col_letter}{total_formula_row_num}"].number_format = '#,##0.00 €'
                                                     # worksheet[f"{total_col_letter}{total_formula_row_num}"].font = Font(bold=True)
                                                     # worksheet[f"{label_col_letter}{total_formula_row_num}"].font = Font(bold=True) # Requires label_col_letter

                                                # Optional: Apply formatting to Minimum Required row label/value if needed
                                                # min_req_row_num = total_formula_row_num + 1
                                                # worksheet[f"{label_col_letter}{min_req_row_num}"].font = Font(bold=True) # Requires label_col_letter
                                                # worksheet[f"{total_col_letter}{min_req_row_num}"].font = Font(bold=True)

                                                sheets_created_count += 1
                                                logging.info(f"Export: Wrote sheet {sanitized_name} with formulas.")

                                            except Exception as write_error:
                                                st.error(f"❌ Erreur écriture onglet/formules pour {supplier} ({sanitized_name}): {write_error}")
                                                logging.error(f"Error writing sheet/formulas {sanitized_name}: {write_error}")
                                        else:
                                             logging.info(f"Export: No items for supplier '{supplier}', skipping sheet.")
                                # End of loop for suppliers
                    except Exception as e_writer:
                        st.error(f"❌ Erreur majeure création fichier Excel : {e_writer}")
                        logging.exception("Error during ExcelWriter context:")

                    # --- Download Button ---
                    if sheets_created_count > 0:
                        output.seek(0)
                        suppliers_str = "multiples" if len(selected_fournisseurs) > 1 else sanitize_sheet_name(selected_fournisseurs[0])
                        timestamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M')
                        filename = f"commande_{suppliers_str}_{timestamp}.xlsx"
                        st.download_button(label=f"📥 Télécharger Commandes ({sheets_created_count} Onglet{'s' if sheets_created_count > 1 else ''})", data=output, file_name=filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                        logging.info(f"Download button created for {sheets_created_count} sheets.")
                    elif formula_ready: # Check if formulas were intended but no data rows existed
                         st.info("ℹ️ Aucune quantité à commander trouvée pour l'exportation.")
                         logging.info("No sheets created (no qty > 0).")

                else: # df_export_all was empty
                    st.info("ℹ️ Aucune quantité à commander globale trouvée pour l'exportation.")
                    logging.info("df_export_all was empty, skipping export.")

            else: # Calculation result was None
                st.error("❌ Le calcul n'a pas pu aboutir.")

        # --- Conditions for no calculation ---
        elif not selected_fournisseurs: st.warning("⚠️ Veuillez sélectionner au moins un fournisseur.")
        elif not semaine_columns and not df_filtered.empty: st.warning("⚠️ Calcul impossible: pas de colonnes ventes valides.")

# --- App footer/initial message ---
elif not uploaded_file:
    st.info("👋 Bienvenue ! Chargez votre fichier Excel.")
