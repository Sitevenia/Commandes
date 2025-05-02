import streamlit as st
import pandas as pd
import numpy as np
import io
import logging
import re
import openpyxl # Ensure openpyxl is imported

# --- Logging Configuration ---
# Setup basic logging (optional but helpful for debugging)
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- Helper Functions ---

def safe_read_excel(uploaded_file, sheet_name, **kwargs):
    """
    Safely reads an Excel sheet, returning None if sheet not found or error occurs.
    Handles BytesIO seeking.
    """
    try:
        # Ensure BytesIO is seekable if passed directly
        if isinstance(uploaded_file, io.BytesIO):
            uploaded_file.seek(0) # Reset buffer position before reading
        return pd.read_excel(uploaded_file, sheet_name=sheet_name, **kwargs)
    except ValueError as e:
        # Specific check for sheet not found error message
        if f"Worksheet named '{sheet_name}' not found" in str(e):
             logging.warning(f"Sheet '{sheet_name}' not found.")
             st.warning(f"⚠️ L'onglet '{sheet_name}' n'a pas été trouvé dans le fichier Excel. Les vérifications associées seront ignorées.")
        else:
             # Other ValueErrors might occur (e.g., unsupported format features)
             logging.error(f"ValueError reading sheet '{sheet_name}': {e}")
             st.error(f"❌ Erreur de valeur lors de la lecture de l'onglet '{sheet_name}': {e}. Vérifiez le format.")
        return None
    except FileNotFoundError: # Although reading from buffer, pandas might raise this internally sometimes
        logging.error(f"FileNotFoundError (internally) reading sheet '{sheet_name}'.")
        st.error(f"❌ Fichier non trouvé (erreur interne pandas) lors de la lecture de l'onglet '{sheet_name}'.")
        return None
    # Catch potential errors from the Excel engine (e.g., bad zip file)
    except Exception as e:
        # More specific exceptions like zipfile.BadZipFile could be caught if needed
        logging.error(f"Unexpected error reading sheet '{sheet_name}': {type(e).__name__} - {e}")
        st.error(f"❌ Erreur inattendue ({type(e).__name__}) lors de la lecture de l'onglet '{sheet_name}': {e}. Le fichier est peut-être corrompu.")
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
            idx_pointer = 0; max_iterations = len(df_calc) * 10; iterations = 0 # Safety break
            while montant_total_actuel < montant_minimum_input and iterations < max_iterations:
                iterations += 1
                if not indices_commandes: # Cannot increase if no items are ordered
                    logging.warning(f"Cannot reach minimum input {montant_minimum_input:.2f}€: no items initially ordered.")
                    break
                current_idx = indices_commandes[idx_pointer % len(indices_commandes)]
                cond = conditionnement.iloc[current_idx]; prix = tarif_achat.iloc[current_idx]

                if cond > 0 and prix > 0: # Only increase if valid conditionnement and price
                    quantite_a_commander[current_idx] += cond; montant_total_actuel += cond * prix
                elif cond <= 0 : # If conditionnement is invalid, remove from pool to avoid infinite loop
                    logging.warning(f"Invalid conditionnement (<=0) for idx {current_idx} during min amount adjustment. Removing.")
                    indices_commandes.pop(idx_pointer % len(indices_commandes))
                    if not indices_commandes: continue # Check again if list is empty after pop
                    idx_pointer -= 1 # Adjust pointer because list size changed
                # Implicitly ignore items with price <= 0 as they don't help reach the minimum
                idx_pointer += 1

            if iterations >= max_iterations and montant_total_actuel < montant_minimum_input:
                 logging.error(f"Failed to reach minimum amount {montant_minimum_input:.2f}€ after {max_iterations} iterations. Reached: {montant_total_actuel:.2f}€.")
                 st.error("L'ajustement automatique pour atteindre le montant minimum a échoué (max iterations). Vérifiez conditionnements/prix.")

        # --- Montant Final ---
        montant_total_final = sum(q * p for q, p in zip(quantite_a_commander, tarif_achat))
        return (quantite_a_commander, ventes_N1, ventes_12_semaines_N1, ventes_12_dernieres_semaines, montant_total_final)

    except KeyError as e: st.error(f"Erreur clé pendant calcul: '{e}'."); logging.error(f"KeyError calc: {e}"); return None
    except ValueError as e: st.error(f"Erreur valeur pendant calcul: {e}"); logging.error(f"ValueError calc: {e}"); return None
    except Exception as e: st.error(f"Erreur inattendue pendant calcul: {e}"); logging.exception("Error calc:"); return None


def sanitize_sheet_name(name):
    """Removes invalid characters for Excel sheet names and truncates."""
    if not isinstance(name, str): name = str(name)
    # Remove specific invalid characters: []:*?/\\<>|"
    sanitized = re.sub(r'[\[\]:*?/\\<>|"]', '_', name)
    # Sheet names cannot start or end with an apostrophe
    if sanitized.startswith("'"): sanitized = "_" + sanitized[1:]
    if sanitized.endswith("'"): sanitized = sanitized[:-1] + "_"
    # Truncate to maximum length (31 characters)
    return sanitized[:31]


# --- Streamlit App Main Logic ---
st.set_page_config(page_title="Forecast App", layout="wide")
st.title("📦 Application de Prévision des Commandes")

uploaded_file = st.file_uploader("📁 Charger le fichier Excel principal", type=["xlsx", "xls"])

# Initialize variables outside the 'if uploaded_file' block
df_full = None
min_order_dict = {}
df_min_commande = None

if uploaded_file:
    try: # Outer try block for initial loading and processing of both sheets
        file_buffer = io.BytesIO(uploaded_file.getvalue())
        st.info("Tentative de lecture de l'onglet 'Tableau final'...")
        logging.info("Attempting to read 'Tableau final' sheet.")

        # --- Read Main Sheet ("Tableau final") ---
        # No inner try needed here as safe_read_excel handles errors gracefully
        df_full = safe_read_excel(file_buffer, sheet_name="Tableau final", header=7)

        if df_full is None:
             # safe_read_excel already showed a warning/error
             st.error("❌ Échec de la lecture de l'onglet principal 'Tableau final'. Impossible de continuer.")
             logging.error("Reading 'Tableau final' failed (safe_read_excel returned None). Stopping.")
             st.stop() # Stop execution if the main sheet cannot be read
        else:
             st.success("✅ Onglet 'Tableau final' lu avec succès.")
             logging.info("'Tableau final' sheet read successfully.")


        # --- Read Minimum Order Sheet ("Minimum de commande") ---
        st.info("Tentative de lecture de l'onglet 'Minimum de commande'...")
        logging.info("Attempting to read 'Minimum de commande' sheet.")
        # safe_read_excel handles seek(0) internally if needed for BytesIO
        df_min_commande = safe_read_excel(file_buffer, sheet_name="Minimum de commande")

        if df_min_commande is not None:
            st.success("✅ Onglet 'Minimum de commande' lu.")
            logging.info("'Minimum de commande' sheet read.")
            # Process Minimum Order Data
            supplier_col_min = "Fournisseur" # Adjust if name is different
            min_amount_col = "Minimum Commande €" # Adjust if name is different
            required_min_cols = [supplier_col_min, min_amount_col]

            if all(col in df_min_commande.columns for col in required_min_cols):
                try:
                    df_min_commande[supplier_col_min] = df_min_commande[supplier_col_min].astype(str).str.strip()
                    df_min_commande[min_amount_col] = pd.to_numeric(df_min_commande[min_amount_col], errors='coerce')
                    min_order_dict = df_min_commande.dropna(subset=[supplier_col_min, min_amount_col])\
                                                .set_index(supplier_col_min)[min_amount_col]\
                                                .to_dict()
                    logging.info(f"Created minimum order dict: {len(min_order_dict)} entries.")
                except Exception as e_min_proc:
                    st.error(f"❌ Erreur lors du traitement des données de 'Minimum de commande': {e_min_proc}")
                    logging.exception("Error processing min order sheet data:")
                    min_order_dict = {} # Reset dict on error
            else:
                missing_min_cols = [col for col in required_min_cols if col not in df_min_commande.columns]
                st.warning(f"⚠️ Colonnes requises ({', '.join(missing_min_cols)}) manquantes dans l'onglet 'Minimum de commande'. La vérification des minimums pourrait être incomplète.")
                logging.warning(f"Missing columns in 'Minimum de commande': {missing_min_cols}")
        # else: safe_read_excel already displayed a warning if sheet wasn't found

    except Exception as e_load:
        # Catch any other unexpected error during the initial loading/reading phase
        st.error(f"❌ Erreur inattendue lors de la lecture du fichier : {e_load}")
        logging.exception("Unexpected error during file loading phase:")
        st.stop()


    # --- Continue with Main Processing only if 'Tableau final' was read successfully ---
    # This check is implicitly handled because we st.stop() if df_full is None
    if df_full is not None: # Technically redundant due to st.stop(), but good for clarity
        try:
            # --- Initial Filtering (Suppliers, Refs) ---
            df = df_full[
                (df_full["Fournisseur"].notna()) & (df_full["Fournisseur"] != "") & (df_full["Fournisseur"] != "#FILTER") &
                (df_full["AF_RefFourniss"].notna()) & (df_full["AF_RefFourniss"] != "")
            ].copy()

            if df.empty:
                 st.warning("Aucune ligne valide trouvée après le filtrage initial (Fournisseur/AF_RefFourniss).")
                 fournisseurs = []
            else:
                fournisseurs = sorted(df["Fournisseur"].unique().tolist())

        except KeyError as e_filter:
            st.error(f"❌ Colonne essentielle '{e_filter}' manquante dans 'Tableau final' pour le filtrage initial.")
            logging.error(f"KeyError during initial filtering: {e_filter}")
            st.stop() # Stop execution if basic filtering fails


        # --- User Selection ---
        selected_fournisseurs = st.multiselect(
            "👤 Sélectionnez le(s) fournisseur(s)",
            options=fournisseurs,
            default=[]
        )

        if selected_fournisseurs:
            df_filtered = df[df["Fournisseur"].isin(selected_fournisseurs)].copy()
        else:
            # Create empty DF with same columns to avoid errors later if no selection
            df_filtered = pd.DataFrame(columns=df.columns)


        # --- Identify Week Columns & Prepare Data ---
        start_col_index = 12 # Index of column N (13th column)
        semaine_columns = []
        if len(df_filtered.columns) > start_col_index:
            potential_week_cols = df_filtered.columns[start_col_index:].tolist()
            # Define columns to exclude from week calculation more robustly
            exclude_columns = [
                "Tarif d'achat", "Conditionnement", "Stock", "Total", "Stock à terme",
                "Ventes N-1", "Ventes 12 semaines identiques N-1", "Ventes 12 dernières semaines",
                "Quantité à commander",
                # Also exclude known non-week columns by name if possible
                "Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article"
            ]
            # Filter potential week columns more carefully
            semaine_columns = [
                col for col in potential_week_cols
                if col not in exclude_columns and pd.api.types.is_numeric_dtype(df_filtered.get(col, pd.Series(dtype=float)).dtype)
                # Optional: Add a check for column name pattern like 'S' followed by number?
                # and (col.startswith('S') and col[1:].isdigit())
            ]

            if not semaine_columns and not df_filtered.empty :
                 st.warning("⚠️ Aucune colonne numérique ressemblant à des ventes hebdomadaires n'a été identifiée après la colonne M.")

            # Ensure essential numeric columns exist and are numeric in the filtered data
            essential_numeric_cols = ["Stock", "Conditionnement", "Tarif d'achat"]
            for col in essential_numeric_cols:
                 if col in df_filtered.columns:
                     df_filtered[col] = pd.to_numeric(df_filtered[col], errors='coerce').fillna(0)
                 elif not df_filtered.empty: # Only error if df is supposed to have data
                     st.error(f"Colonne essentielle '{col}' manquante dans les données filtrées pour les fournisseurs sélectionnés.")
                     st.stop() # Stop if essential calculation columns are missing

        elif not df_filtered.empty:
            st.warning("Le fichier ne semble pas contenir de colonnes après l'index 12 (colonne M) pour les données de ventes.")


        # --- Calculation Parameters ---
        col1, col2 = st.columns(2)
        with col1:
            duree_semaines = st.number_input("⏳ Durée couverture (semaines)", value=4, min_value=1, step=1, key="duree")
        with col2:
            montant_minimum_input_val = st.number_input(
                "💶 Montant minimum global (€)", value=0.0, min_value=0.0, step=50.0, format="%.2f", key="montant_min",
                help="Montant minimum global utilisé pour tenter d'ajuster les quantités. Une alerte séparée apparaîtra si le minimum spécifique d'un fournisseur unique n'est pas atteint."
                )

        # --- Execute Calculation ---
        if not df_filtered.empty and semaine_columns:
            st.info("🚀 Lancement du calcul des quantités...")
            result = calculer_quantite_a_commander(
                df_filtered,
                semaine_columns,
                montant_minimum_input_val, # Pass the user input value
                duree_semaines
            )

            if result is not None:
                st.success("✅ Calculs terminés.")
                # --- Unpack results and add calculated columns ---
                (quantite_calculée, ventes_N1_calc, ventes_12_N1_calc,
                 ventes_12_last_calc, montant_total_calc) = result

                df_filtered.loc[:, "Quantité à commander"] = quantite_calculée
                df_filtered.loc[:, "Ventes N-1"] = ventes_N1_calc
                df_filtered.loc[:, "Ventes 12 semaines identiques N-1"] = ventes_12_N1_calc
                df_filtered.loc[:, "Ventes 12 dernières semaines"] = ventes_12_last_calc
                # Ensure recalculation of Total and Stock à terme based on final Quantité à commander
                df_filtered.loc[:, "Total"] = df_filtered["Tarif d'achat"] * df_filtered["Quantité à commander"]
                df_filtered.loc[:, "Stock à terme"] = df_filtered["Stock"] + df_filtered["Quantité à commander"]

                st.metric(label="💰 Montant total GLOBAL calculé", value=f"{montant_total_calc:.2f} €")

                # --- MINIMUM WARNING (for single supplier selection only) ---
                if len(selected_fournisseurs) == 1:
                    selected_supplier = selected_fournisseurs[0]
                    if selected_supplier in min_order_dict:
                        required_minimum = min_order_dict[selected_supplier]
                        # Calculate the actual total for THIS supplier from the results DataFrame
                        supplier_actual_total = df_filtered[df_filtered["Fournisseur"] == selected_supplier]["Total"].sum()
                        if required_minimum > 0 and supplier_actual_total < required_minimum:
                            diff = required_minimum - supplier_actual_total
                            st.warning(
                                f"⚠️ **Minimum Non Atteint (Fournisseur: {selected_supplier})**\n"
                                f"Montant Calculé pour ce fournisseur: **{supplier_actual_total:.2f} €**\n"
                                f"Minimum Requis (fichier Excel): **{required_minimum:.2f} €** (Manque: {diff:.2f} €)\n\n"
                                f"➡️ **Suggestion:** Pour tenter d'atteindre ce minimum spécifique, vous pouvez modifier la valeur du champ 'Montant minimum global (€)' ci-dessus à **{required_minimum:.2f}** (ou plus) et relancer le calcul."
                            )

                # --- Display Results Table (Combined) ---
                st.subheader("📊 Résultats Combinés (Tous Fournisseurs Sélectionnés)")
                # Define columns for display, ensuring Fournisseur is included here
                required_display_columns = ["Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article", "Stock"]
                display_columns_base = required_display_columns + [
                    "Ventes N-1", "Ventes 12 semaines identiques N-1", "Ventes 12 dernières semaines",
                    "Conditionnement", "Quantité à commander", "Stock à terme",
                    "Tarif d'achat", "Total"
                ]
                display_columns = [col for col in display_columns_base if col in df_filtered.columns]
                missing_display_columns = [col for col in required_display_columns if col not in df_filtered.columns]

                if missing_display_columns:
                    st.error(f"❌ Colonnes manquantes pour l'affichage des résultats combinés : {', '.join(missing_display_columns)}")
                else:
                    # Apply formatting using style
                    st.dataframe(df_filtered[display_columns].style.format({
                        "Tarif d'achat": "{:,.2f}€", # Added thousands separator
                        "Total": "{:,.2f}€",
                        "Ventes N-1": "{:,.0f}",
                        "Ventes 12 semaines identiques N-1": "{:,.0f}",
                        "Ventes 12 dernières semaines": "{:,.0f}",
                        "Stock": "{:,.0f}",
                        "Conditionnement": "{:,.0f}",
                        "Quantité à commander": "{:,.0f}",
                        "Stock à terme": "{:,.0f}"
                    }, na_rep="-", thousands=",")) # Use comma as thousands separator


                # --- EXPORT LOGIC (Multi-Sheet) ---
                st.subheader("⬇️ Exportation Excel par Fournisseur")
                # Filter *once* for items with quantity > 0 across *all* selected suppliers
                df_export_all = df_filtered[df_filtered["Quantité à commander"] > 0].copy()

                if not df_export_all.empty:
                    output = io.BytesIO()
                    sheets_created_count = 0 # Counter for actual sheets written
                    try:
                        # Use ExcelWriter context manager
                        with pd.ExcelWriter(output, engine="openpyxl") as writer:
                            logging.info(f"Starting export loop for suppliers: {selected_fournisseurs}")
                            # Define columns for export sheets - IMPORTANT: Exclude 'Fournisseur' if not needed on individual sheets
                            export_columns = [col for col in display_columns if col != 'Fournisseur']

                            for supplier in selected_fournisseurs:
                                logging.info(f"Processing export for supplier: {supplier}")
                                # Filter for the current supplier's items with quantity > 0
                                df_supplier_export = df_export_all[df_export_all["Fournisseur"] == supplier]

                                if not df_supplier_export.empty:
                                    logging.info(f"Found {len(df_supplier_export)} items to order for {supplier}.")
                                    # Prepare the main data part for the sheet
                                    df_supplier_sheet_data = df_supplier_export[export_columns].copy()

                                    # --- Create Summary Rows ---
                                    supplier_total = df_supplier_sheet_data["Total"].sum()
                                    required_minimum = min_order_dict.get(supplier, 0) # Default to 0 if not found
                                    min_formatted = f"{required_minimum:,.2f} €" if required_minimum > 0 else "N/A" # Format minimum

                                    # Determine columns for labels and values in summary rows
                                    label_col = "Désignation Article" if "Désignation Article" in export_columns else export_columns[1] # Fallback: Use 2nd col (usually Ref Article)
                                    value_col = "Total" if "Total" in export_columns else export_columns[-1] # Fallback: Use last col

                                    # Build Total Row DataFrame
                                    total_row_dict = {col: "" for col in export_columns}
                                    total_row_dict[label_col] = "TOTAL COMMANDE"
                                    total_row_dict[value_col] = supplier_total # Store numeric value for potential formatting
                                    total_row_df = pd.DataFrame([total_row_dict])

                                    # Build Minimum Row DataFrame
                                    min_row_dict = {col: "" for col in export_columns}
                                    min_row_dict[label_col] = "Minimum Requis"
                                    min_row_dict[value_col] = min_formatted # Store formatted string
                                    min_row_df = pd.DataFrame([min_row_dict])

                                    # Concatenate data + total row + minimum row
                                    df_sheet = pd.concat([df_supplier_sheet_data, total_row_df, min_row_df], ignore_index=True)

                                    # Sanitize sheet name
                                    sanitized_name = sanitize_sheet_name(supplier)
                                    logging.info(f"Using sanitized sheet name: {sanitized_name}")

                                    # Write to Excel sheet within the context manager
                                    try:
                                        df_sheet.to_excel(writer, sheet_name=sanitized_name, index=False)
                                        # --- Optional: Apply formatting using openpyxl ---
                                        # worksheet = writer.sheets[sanitized_name]
                                        # from openpyxl.styles import Font, Alignment, NumberFormat
                                        # # Example: Make summary rows bold and format total currency
                                        # bold_font = Font(bold=True)
                                        # last_row_idx = len(df_sheet) # Index of Minimum row (1-based for openpyxl)
                                        # total_row_idx = last_row_idx - 1
                                        # # Get column letter (requires column index + 1)
                                        # try:
                                        #      value_col_letter = openpyxl.utils.get_column_letter(export_columns.index(value_col) + 1)
                                        #      label_col_letter = openpyxl.utils.get_column_letter(export_columns.index(label_col) + 1)
                                        #
                                        #      # Apply bold to labels and values in summary rows
                                        #      for row_idx in [total_row_idx, last_row_idx]:
                                        #           worksheet[f"{label_col_letter}{row_idx}"].font = bold_font
                                        #           worksheet[f"{value_col_letter}{row_idx}"].font = bold_font
                                        #
                                        #      # Format the actual total value as currency (won't affect the Min Required text)
                                        #      total_cell = worksheet[f"{value_col_letter}{total_row_idx}"]
                                        #      total_cell.number_format = '#,##0.00 €' # Apply currency format
                                        # except (ValueError, IndexError) as fmt_err:
                                        #      logging.warning(f"Could not apply Excel formatting for sheet {sanitized_name}: {fmt_err}")

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
                            timestamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M')
                            filename = f"commande_{suppliers_str}_{timestamp}.xlsx"

                            st.download_button(
                                label=f"📥 Télécharger Commandes ({sheets_created_count} Onglet{'s' if sheets_created_count > 1 else ''})", # Dynamic label
                                data=output,
                                file_name=filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                            logging.info(f"Download button created for {sheets_created_count} sheets.")
                        else:
                            st.info("ℹ️ Aucune quantité à commander trouvée pour l'exportation pour les fournisseurs sélectionnés.")
                            logging.info("No sheets were created as no items had quantity > 0 for any selected supplier.")

                    # Catch errors during the ExcelWriter process itself
                    except Exception as e_writer:
                        st.error(f"❌ Erreur majeure lors de la création du fichier Excel : {e_writer}")
                        logging.exception("Error during ExcelWriter context or processing:")

                else:
                    st.info("ℹ️ Aucune quantité à commander globale trouvée pour l'exportation.")
                    logging.info("df_export_all was empty, skipping export.")

            # --- Error handling for calculation failure ---
            else:
                st.error("❌ Le calcul des quantités n'a pas pu aboutir. Vérifiez les messages d'erreur précédents.")
                logging.error("Calculation result was None.")

        # --- Conditions for no calculation / no selection / no columns ---
        elif not selected_fournisseurs:
            st.warning("⚠️ Veuillez sélectionner au moins un fournisseur pour lancer le calcul.")
        elif not semaine_columns and not df_filtered.empty: # Filtered DF exists but no week columns found
            st.warning("⚠️ Calcul impossible: aucune colonne de ventes hebdomadaires valide n'a été identifiée ou les données sont incomplètes.")
        # Implicitly handles df_filtered.empty case (no suppliers selected or no valid data)

# --- App footer or initial message ---
elif not uploaded_file:
    st.info("👋 Bienvenue ! Chargez votre fichier Excel contenant les données de ventes et de stock pour commencer.")

# --- Final catch-all (optional, might hide more specific errors) ---
# except Exception as e:
#     st.error(f"❌ Une erreur générale et imprévue est survenue: {e}")
#     logging.exception("Unhandled exception occurred in the main app flow:")
