import streamlit as st
import pandas as pd
import numpy as np
import io
import logging
import re
import openpyxl # Required for engine and direct manipulation
from openpyxl.utils import get_column_letter # Utility to get column letters
# from openpyxl.styles import Font # Uncomment if applying bold font formatting

# --- Logging Configuration ---
# Setup basic logging (INFO level is usually good, DEBUG for more detail)
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
        # Specify the engine explicitly, especially if dealing with various formats or macros
        engine = 'openpyxl' if str(uploaded_file.name).lower().endswith('.xlsx') else None # Use default for .xls
        return pd.read_excel(uploaded_file, sheet_name=sheet_name, engine=engine, **kwargs)
    except ValueError as e:
        # Specific check for sheet not found error message
        # Pandas error messages might vary slightly across versions
        if f"Worksheet named '{sheet_name}' not found" in str(e) or f"'{sheet_name}' not found" in str(e):
             logging.warning(f"Sheet '{sheet_name}' not found in the uploaded file.")
             st.warning(f"⚠️ L'onglet '{sheet_name}' n'a pas été trouvé dans le fichier Excel. Les fonctionnalités associées pourraient être absentes.")
        else:
             # Other ValueErrors might occur (e.g., unsupported format features)
             logging.error(f"ValueError reading sheet '{sheet_name}': {e}")
             st.error(f"❌ Erreur de valeur lors de la lecture de l'onglet '{sheet_name}': {e}. Vérifiez le format ou le contenu de l'onglet.")
        return None
    except FileNotFoundError: # Although reading from buffer, pandas might raise this internally
        logging.error(f"FileNotFoundError (internally) reading sheet '{sheet_name}'.")
        st.error(f"❌ Fichier non trouvé (erreur interne pandas) lors de la lecture de l'onglet '{sheet_name}'.")
        return None
    # Catch potential errors from the Excel engine (e.g., bad zip file for xlsx)
    except Exception as e:
        # Example: Catching bad zip file error specifically
        if "zip file" in str(e).lower():
             logging.error(f"Error reading sheet '{sheet_name}': Bad zip file - {e}")
             st.error(f"❌ Erreur lors de la lecture de l'onglet '{sheet_name}': Le fichier Excel (.xlsx) semble corrompu (erreur zip).")
        else:
            logging.error(f"Unexpected error reading sheet '{sheet_name}': {type(e).__name__} - {e}")
            st.error(f"❌ Erreur inattendue ({type(e).__name__}) lors de la lecture de l'onglet '{sheet_name}': {e}. Le fichier est peut-être corrompu ou dans un format inattendu.")
        return None

def calculer_quantite_a_commander(df, semaine_columns, montant_minimum_input, duree_semaines):
    """
    Calcule la quantité à commander pour chaque produit en fonction des ventes passées,
    du stock actuel, du conditionnement et d'un montant minimum de commande (fourni en entrée).
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
            elif q > 0: quantite_a_commander[i] = 0 # Conditionnement invalide ou nul, on ne commande pas
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
            # Optional: Sort indices by price * conditionnement descending?
            # indices_commandes.sort(key=lambda i: tarif_achat.iloc[i] * conditionnement.iloc[i] if conditionnement.iloc[i]>0 else 0, reverse=True)

            idx_pointer = 0
            max_iterations = len(df_calc) * 10 # Safety break
            iterations = 0

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
                    # Need to check if list is empty *after* removing the item
                    if not indices_commandes: continue # Go to next iteration check (will break if empty)
                    # Adjust pointer because list size changed and we removed the item at current pointer mod len
                    idx_pointer -= 1 # Stay on the "next" item logically after removal
                # Implicitly ignore items with price <= 0 as they don't help reach the minimum

                idx_pointer += 1 # Move to next candidate index

            if iterations >= max_iterations and montant_total_actuel < montant_minimum_input:
                 logging.error(f"Failed to reach minimum amount {montant_minimum_input:.2f}€ after {max_iterations} iterations. Reached: {montant_total_actuel:.2f}€.")
                 st.error("L'ajustement automatique pour atteindre le montant minimum global a échoué (max iterations dépassées). Vérifiez les conditionnements et tarifs des articles initialement commandés.")

        # --- Montant Final ---
        # Recalculate final total based on the potentially adjusted quantite_a_commander list
        montant_total_final = sum(q * p for q, p in zip(quantite_a_commander, tarif_achat))

        return (quantite_a_commander,
                ventes_N1,
                ventes_12_semaines_N1,
                ventes_12_dernieres_semaines,
                montant_total_final) # Return the final calculated amount

    except KeyError as e:
        st.error(f"Erreur de clé pendant le calcul: La colonne '{e}' est introuvable.")
        logging.error(f"KeyError during calculation: {e}")
        return None
    except ValueError as e:
         st.error(f"Erreur de valeur pendant le calcul: Problème avec les données numériques - {e}")
         logging.error(f"ValueError during calculation: {e}")
         return None
    except Exception as e:
        st.error(f"Erreur inattendue pendant le calcul des quantités: {e}")
        logging.exception("Unexpected error during quantity calculation:") # Log full traceback
        return None

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

uploaded_file = st.file_uploader("📁 Charger le fichier Excel principal", type=["xlsx", "xls"], key="fileUploader")

# Initialize variables outside the 'if uploaded_file' block
df_full = None
min_order_dict = {}
df_min_commande = None

if uploaded_file:
    # Add a button to clear the uploaded file and reset state if needed
    # if st.button("Réinitialiser / Charger un nouveau fichier"):
    #     st.cache_data.clear() # Clear pandas cache if used
    #     st.experimental_rerun() # Rerun the script

    try: # Outer try block for initial loading and processing of both sheets
        file_buffer = io.BytesIO(uploaded_file.getvalue())
        st.info("Lecture onglet 'Tableau final'...")
        logging.info(f"Attempting to read 'Tableau final' sheet from file: {uploaded_file.name}")

        # --- Read Main Sheet ("Tableau final") ---
        df_full = safe_read_excel(file_buffer, sheet_name="Tableau final", header=7)

        if df_full is None:
             # Error/Warning already displayed by safe_read_excel
             st.error("❌ Échec de la lecture de l'onglet principal 'Tableau final'. Vérifiez le nom de l'onglet et l'intégrité du fichier. Impossible de continuer.")
             logging.error("Reading 'Tableau final' failed (safe_read_excel returned None). Stopping.")
             st.stop() # Stop execution if the main sheet cannot be read
        else:
             st.success("✅ Onglet 'Tableau final' lu avec succès.")
             logging.info("'Tableau final' sheet read successfully.")


        # --- Read Minimum Order Sheet ("Minimum de commande") ---
        st.info("Lecture onglet 'Minimum de commande'...")
        logging.info("Attempting to read 'Minimum de commande' sheet.")
        df_min_commande = safe_read_excel(file_buffer, sheet_name="Minimum de commande")

        if df_min_commande is not None:
            st.success("✅ Onglet 'Minimum de commande' lu.")
            logging.info("'Minimum de commande' sheet read.")
            # Process Minimum Order Data
            supplier_col_min = "Fournisseur" # Adjust if name is different in your file
            min_amount_col = "Minimum de Commande" # Adjust if name is different
            required_min_cols = [supplier_col_min, min_amount_col]

            if all(col in df_min_commande.columns for col in required_min_cols):
                try:
                    # Ensure supplier name is string and clean, amount is numeric
                    df_min_commande[supplier_col_min] = df_min_commande[supplier_col_min].astype(str).str.strip()
                    df_min_commande[min_amount_col] = pd.to_numeric(df_min_commande[min_amount_col], errors='coerce')
                    # Create dictionary, dropping rows with missing supplier or non-numeric minimum
                    min_order_dict = df_min_commande.dropna(subset=[supplier_col_min, min_amount_col])\
                                                .set_index(supplier_col_min)[min_amount_col]\
                                                .to_dict()
                    logging.info(f"Successfully created minimum order dictionary with {len(min_order_dict)} entries.")
                except Exception as e_min_proc:
                    st.error(f"❌ Erreur lors du traitement des données de l'onglet 'Minimum de commande': {e_min_proc}")
                    logging.exception("Error processing minimum order sheet data:")
                    min_order_dict = {} # Reset dict on error
            else:
                missing_min_cols = [col for col in required_min_cols if col not in df_min_commande.columns]
                st.warning(f"⚠️ Colonnes requises ({', '.join(missing_min_cols)}) manquantes dans l'onglet 'Minimum de commande'. La vérification des minimums spécifiques par fournisseur sera désactivée.")
                logging.warning(f"Missing required columns in 'Minimum de commande': {missing_min_cols}")
        # else: safe_read_excel already displayed a warning if sheet wasn't found

    except Exception as e_load:
        # Catch any other unexpected error during the initial loading/reading phase
        st.error(f"❌ Erreur inattendue lors de la lecture du fichier Excel : {e_load}")
        logging.exception("Unexpected error during file loading phase:")
        st.stop()


    # --- Continue with Main Processing only if 'Tableau final' was read successfully ---
    if df_full is not None: # Main data is loaded
        try:
            # --- Initial Filtering (Suppliers, Refs) ---
            # Ensure required columns for filtering exist before attempting to filter
            filter_cols = ["Fournisseur", "AF_RefFourniss"]
            if not all(col in df_full.columns for col in filter_cols):
                 st.error(f"❌ Colonnes nécessaires pour le filtrage initial ({', '.join(filter_cols)}) manquantes dans 'Tableau final'.")
                 st.stop()

            df = df_full[
                (df_full["Fournisseur"].notna()) & (df_full["Fournisseur"] != "") & (df_full["Fournisseur"] != "#FILTER") &
                (df_full["AF_RefFourniss"].notna()) & (df_full["AF_RefFourniss"] != "")
            ].copy()

            fournisseurs = sorted(df["Fournisseur"].unique().tolist()) if not df.empty else []
            if df.empty and not df_full.empty: # df_full had rows, but none passed the filter
                st.warning("Aucune ligne valide trouvée après le filtrage initial (Fournisseur renseigné et non '#FILTER', AF_RefFourniss renseignée).")
            elif not df_full.empty and not fournisseurs: # Should not happen if df is not empty, but as safeguard
                 st.warning("Impossible d'extraire la liste des fournisseurs.")


        except KeyError as e_filter:
            # This catch might be redundant now due to the check above, but kept for safety
            st.error(f"❌ Erreur de clé lors du filtrage initial: Colonne '{e_filter}' non trouvée.")
            logging.error(f"KeyError during initial filtering: {e_filter}")
            st.stop() # Stop execution if basic filtering setup fails
        except Exception as e_filter_other:
             st.error(f"❌ Erreur inattendue lors du filtrage initial : {e_filter_other}")
             logging.exception("Unexpected error during initial filtering:")
             st.stop()


        # --- User Selection ---
        st.subheader("1. Sélection des Fournisseurs et Paramètres")
        selected_fournisseurs = st.multiselect(
            "👤 Sélectionnez le(s) fournisseur(s) pour le calcul",
            options=fournisseurs,
            default=[],
            key="supplier_select"
        )

        if selected_fournisseurs:
            df_filtered = df[df["Fournisseur"].isin(selected_fournisseurs)].copy()
            st.info(f"{len(df_filtered)} articles trouvés pour le(s) fournisseur(s) sélectionné(s).")
        else:
            # Create empty DF with same columns to avoid errors later if no selection
            df_filtered = pd.DataFrame(columns=df.columns)


        # --- Identify Week Columns & Prepare Data ---
        start_col_index = 12 # Index of column N (13th column) is 'M'
        semaine_columns = []
        if len(df_filtered.columns) > start_col_index:
            potential_week_cols = df_filtered.columns[start_col_index:].tolist()
            # Define columns to exclude from week calculation more robustly
            # Ensure these names exactly match your DataFrame columns
            exclude_columns = [
                "Tarif d'achat", "Conditionnement", "Stock", "Total", "Stock à terme",
                "Ventes N-1", "Ventes 12 semaines identiques N-1", "Ventes 12 dernières semaines",
                "Quantité à commander",
                # Also exclude known non-week columns by name
                "Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article"
            ]
            # Filter potential week columns more carefully
            semaine_columns = [
                col for col in potential_week_cols
                if col not in exclude_columns and pd.api.types.is_numeric_dtype(df_filtered.get(col, pd.Series(dtype=float)).dtype)
                # Optional: Add a check for column name pattern? e.g., starts with 'S' or 'W' and digits?
                # and (col.startswith(('S','W')) and col[1:].isdigit())
            ]
            logging.info(f"Identified {len(semaine_columns)} potential week columns starting after index {start_col_index}.")

            if not semaine_columns and not df_filtered.empty :
                 st.warning("⚠️ Aucune colonne numérique interprétée comme ventes hebdomadaires n'a été trouvée après la colonne M.")

            # Ensure essential numeric columns exist and are numeric in the filtered data
            essential_numeric_cols = ["Stock", "Conditionnement", "Tarif d'achat"]
            missing_essential = False
            for col in essential_numeric_cols:
                 if col in df_filtered.columns:
                     # Convert to numeric, coercing errors to NaN, then fill NaN with 0
                     df_filtered[col] = pd.to_numeric(df_filtered[col], errors='coerce').fillna(0)
                 elif not df_filtered.empty: # Only error if df is supposed to have data
                     st.error(f"Colonne essentielle '{col}' manquante dans les données pour le(s) fournisseur(s) sélectionné(s). Le calcul ne peut pas continuer.")
                     missing_essential = True
            if missing_essential:
                 st.stop() # Stop if essential calculation columns are missing

        elif not df_filtered.empty: # Suppliers selected, data exists, but file too short?
            st.warning("Le fichier ne semble pas contenir de colonnes après l'index 12 (colonne M) pour les données de ventes.")


        # --- Calculation Parameters ---
        col1, col2 = st.columns(2)
        with col1:
            duree_semaines = st.number_input("⏳ Durée couverture souhaitée (semaines)", value=4, min_value=1, step=1, key="duree", help="Nombre de semaines de ventes futures estimées que la commande doit couvrir.")
        with col2:
            montant_minimum_input_val = st.number_input(
                "💶 Montant minimum global (€)", value=0.0, min_value=0.0, step=50.0, format="%.2f", key="montant_min",
                help="Montant minimum global utilisé pour tenter d'ajuster les quantités à la hausse. Si un seul fournisseur est sélectionné, une alerte séparée apparaîtra si son minimum spécifique (lu depuis l'onglet 'Minimum de commande') n'est pas atteint."
                )

        # --- Execute Calculation ---
        st.subheader("2. Lancement du Calcul")
        # Add a button to trigger calculation for clarity
        if st.button("🚀 Calculer les Quantités à Commander", key="calculate_button"):
            if not df_filtered.empty and semaine_columns:
                st.info("Calcul en cours...")
                logging.info("Starting quantity calculation...")
                result = calculer_quantite_a_commander(
                    df_filtered,
                    semaine_columns,
                    montant_minimum_input_val, # Pass the user input value
                    duree_semaines
                )

                if result is not None:
                    st.success("✅ Calculs terminés.")
                    logging.info("Quantity calculation finished successfully.")
                    # --- Unpack results and add/update calculated columns ---
                    (quantite_calculée, ventes_N1_calc, ventes_12_N1_calc,
                     ventes_12_last_calc, montant_total_calc) = result

                    # Assign results safely using .loc
                    df_filtered.loc[:, "Quantité à commander"] = quantite_calculée
                    df_filtered.loc[:, "Ventes N-1"] = ventes_N1_calc
                    df_filtered.loc[:, "Ventes 12 semaines identiques N-1"] = ventes_12_N1_calc
                    df_filtered.loc[:, "Ventes 12 dernières semaines"] = ventes_12_last_calc
                    # Ensure Tariff is numeric *before* final Total calculation
                    df_filtered.loc[:, "Tarif d'achat"] = pd.to_numeric(df_filtered["Tarif d'achat"], errors='coerce').fillna(0)
                    # Recalculate Total based on the *final* Quantité à commander
                    df_filtered.loc[:, "Total"] = df_filtered["Tarif d'achat"] * df_filtered["Quantité à commander"]
                    # Recalculate Stock à terme
                    df_filtered.loc[:, "Stock à terme"] = df_filtered["Stock"] + df_filtered["Quantité à commander"]

                    # Store results in session state to persist after rerun (e.g., for display/export)
                    st.session_state.calculation_result_df = df_filtered
                    st.session_state.montant_total_calc = montant_total_calc
                    st.session_state.min_order_dict = min_order_dict # Store this too if needed for display
                    st.session_state.selected_fournisseurs_calc = selected_fournisseurs # Store suppliers used for this calc

                    # Force rerun to display results section
                    st.experimental_rerun()

                else: # Calculation function returned None
                    st.error("❌ Le calcul des quantités n'a pas pu aboutir. Vérifiez les messages d'erreur précédents et les données d'entrée.")
                    logging.error("Calculation function returned None.")
                    # Clear previous results if calculation fails?
                    if 'calculation_result_df' in st.session_state:
                        del st.session_state.calculation_result_df


            # --- Conditions for no calculation ---
            elif not selected_fournisseurs:
                st.warning("⚠️ Veuillez sélectionner au moins un fournisseur avant de lancer le calcul.")
            elif df_filtered.empty and selected_fournisseurs:
                 st.warning("⚠️ Aucun article trouvé pour le(s) fournisseur(s) sélectionné(s) après filtrage initial.")
            elif not semaine_columns and not df_filtered.empty:
                st.warning("⚠️ Calcul impossible: aucune colonne de ventes hebdomadaires valide n'a été identifiée pour les articles sélectionnés.")


        # --- Display Results (if calculation was run successfully) ---
        st.subheader("3. Résultats du Calcul")
        if 'calculation_result_df' in st.session_state:
            df_results = st.session_state.calculation_result_df
            montant_total_display = st.session_state.montant_total_calc
            min_order_dict_display = st.session_state.min_order_dict
            suppliers_displayed = st.session_state.selected_fournisseurs_calc

            st.metric(label="💰 Montant total GLOBAL calculé pour la sélection", value=f"{montant_total_display:,.2f} €")

            # --- MINIMUM WARNING (Single Supplier Check) ---
            if len(suppliers_displayed) == 1:
                selected_supplier = suppliers_displayed[0]
                if selected_supplier in min_order_dict_display:
                    required_minimum = min_order_dict_display[selected_supplier]
                    # Use the total calculated for the results dataframe (should match metric)
                    supplier_actual_total = df_results["Total"].sum() # Already filtered for the single supplier
                    if required_minimum > 0 and supplier_actual_total < required_minimum:
                        diff = required_minimum - supplier_actual_total
                        st.warning(
                            f"⚠️ **Minimum Non Atteint (Fournisseur: {selected_supplier})**\n"
                            f"Montant Calculé pour ce fournisseur: **{supplier_actual_total:,.2f} €**\n"
                            f"Minimum Requis (fichier Excel): **{required_minimum:,.2f} €** (Manque: {diff:,.2f} €)\n\n"
                            f"➡️ **Suggestion:** Pour tenter d'atteindre ce minimum spécifique, vous pouvez modifier la valeur du champ 'Montant minimum global (€)' ci-dessus (section 1) à **{required_minimum:.2f}** (ou plus) et relancer le calcul."
                        )

            # --- Display Results Table (Combined) ---
            st.markdown("#### Résultats Détaillés par Article")
            # Define columns for display, ensuring Fournisseur is included
            required_display_columns = ["Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article", "Stock"]
            display_columns_base = required_display_columns + [
                "Ventes N-1", "Ventes 12 semaines identiques N-1", "Ventes 12 dernières semaines",
                "Conditionnement", "Quantité à commander", "Stock à terme",
                "Tarif d'achat", "Total"
            ]
            # Ensure columns exist in the result df before trying to display
            display_columns = [col for col in display_columns_base if col in df_results.columns]
            missing_display_cols_for_table = [col for col in required_display_columns if col not in df_results.columns]

            if missing_display_cols_for_table:
                st.error(f"❌ Impossible d'afficher les résultats détaillés: Colonnes manquantes ({', '.join(missing_display_cols_for_table)}).")
            else:
                # Apply formatting using style
                st.dataframe(df_results[display_columns].style.format({
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


            # --- EXPORT LOGIC (Multi-Sheet with Formulas) ---
            st.markdown("#### Exportation des Commandes")
            # Filter *once* for items with quantity > 0 from the results dataframe
            df_export_all = df_results[df_results["Quantité à commander"] > 0].copy()

            if not df_export_all.empty:
                output_export = io.BytesIO()
                sheets_created_count_export = 0
                try:
                    # Use ExcelWriter context manager
                    with pd.ExcelWriter(output_export, engine="openpyxl") as writer:
                        logging.info(f"Export: Starting for suppliers: {suppliers_displayed}")

                        # --- Define column names (CRITICAL: Must match DataFrame columns) ---
                        qty_col_name = "Quantité à commander"
                        price_col_name = "Tarif d'achat"
                        total_col_name = "Total"
                        # Define columns for export sheets (using display_columns ensures they exist in df_results)
                        export_columns = [col for col in display_columns if col != 'Fournisseur'] # Exclude 'Fournisseur' from individual sheets

                        # --- Verify essential columns exist and get letters ---
                        formula_ready = False
                        if not all(c in export_columns for c in [qty_col_name, price_col_name, total_col_name]):
                            st.error(f"❌ Export Error: Columns '{qty_col_name}', '{price_col_name}', or '{total_col_name}' not found in calculated results for export.")
                            logging.error(f"Export Error: Essential formula columns missing from {export_columns}")
                        else:
                            try:
                                # Get 0-based indices relative to the export_columns list
                                qty_col_idx = export_columns.index(qty_col_name)
                                price_col_idx = export_columns.index(price_col_name)
                                total_col_idx = export_columns.index(total_col_name)
                                # Convert to 1-based Excel column letters
                                qty_col_letter = get_column_letter(qty_col_idx + 1)
                                price_col_letter = get_column_letter(price_col_idx + 1)
                                total_col_letter = get_column_letter(total_col_idx + 1)
                                formula_ready = True
                                logging.info(f"Export: Formula cols identified: Qty={qty_col_letter}, Price={price_col_letter}, Total={total_col_letter}")
                            except ValueError as e:
                                st.error(f"❌ Export Error: Could not find index for formula columns: {e}")
                                logging.error(f"Export Error: Could not get column index: {e}")
                            except Exception as e_idx:
                                 st.error(f"❌ Export Error: Unexpected error getting column indices: {e_idx}")
                                 logging.exception("Export Error: Getting column indices failed.")


                        # --- Proceed only if columns for formula are identified ---
                        if formula_ready:
                            for supplier_export in suppliers_displayed: # Iterate through suppliers used for calculation
                                logging.info(f"Export: Processing supplier sheet for {supplier_export}")
                                df_supplier_export_filtered = df_export_all[df_export_all["Fournisseur"] == supplier_export].copy() # Use copy

                                if not df_supplier_export_filtered.empty:
                                    # Select only the export columns for the sheet data
                                    df_supplier_sheet_data = df_supplier_export_filtered[export_columns].copy()
                                    num_data_rows = len(df_supplier_sheet_data)
                                    logging.info(f"Export: {supplier_export} - Found {num_data_rows} data rows.")

                                    # --- Summary Rows Prep ---
                                    supplier_total_val = df_supplier_sheet_data[total_col_name].sum() # Calculated value (placeholder for SUM formula cell)
                                    required_minimum_export = min_order_dict_display.get(supplier_export, 0)
                                    min_formatted_export = f"{required_minimum_export:,.2f} €" if required_minimum_export > 0 else "N/A"
                                    # Determine label column robustly
                                    if "Désignation Article" in export_columns: label_col = "Désignation Article"
                                    elif "Référence Article" in export_columns: label_col = "Référence Article"
                                    else: label_col = export_columns[1] # Fallback to 2nd column

                                    # Create DataFrames for summary rows
                                    total_row_dict = {col: "" for col in export_columns}; total_row_dict[label_col] = "TOTAL COMMANDE"; total_row_dict[total_col_name] = supplier_total_val # Placeholder value
                                    min_row_dict = {col: "" for col in export_columns}; min_row_dict[label_col] = "Minimum Requis"; min_row_dict[total_col_name] = min_formatted_export
                                    total_row_df = pd.DataFrame([total_row_dict]); min_row_df = pd.DataFrame([min_row_dict])

                                    # Concatenate data + summary rows
                                    df_sheet = pd.concat([df_supplier_sheet_data, total_row_df, min_row_df], ignore_index=True)

                                    sanitized_name = sanitize_sheet_name(supplier_export)
                                    try:
                                        # --- Step 1: Write DataFrame to Excel ---
                                        logging.debug(f"Export: Writing df_sheet to sheet '{sanitized_name}'")
                                        df_sheet.to_excel(writer, sheet_name=sanitized_name, index=False)

                                        # --- Step 2: Get the openpyxl worksheet object ---
                                        worksheet = writer.sheets[sanitized_name]
                                        logging.debug(f"Export: Got worksheet object for '{sanitized_name}'")

                                        # --- Step 3: Overwrite 'Total' cells with FORMULAS ---
                                        logging.debug(f"Export: Applying formulas to rows 2 to {num_data_rows + 1} in col {total_col_letter}")
                                        # Excel rows are 1-based. Header is 1. Data starts row 2.
                                        for excel_row_num in range(2, num_data_rows + 2): # +2 because range end is exclusive
                                            formula_str = f"={qty_col_letter}{excel_row_num}*{price_col_letter}{excel_row_num}"
                                            cell = worksheet[f"{total_col_letter}{excel_row_num}"]
                                            cell.value = formula_str # Assign formula string
                                            cell.number_format = '#,##0.00 €' # Apply number format

                                        # --- Step 4: Apply SUM formula to the grand total row ---
                                        total_formula_row_num = num_data_rows + 2 # Excel row number for "TOTAL COMMANDE"
                                        logging.debug(f"Export: Applying SUM formula to row {total_formula_row_num} in col {total_col_letter}")
                                        if num_data_rows > 0: # Only add sum if there was data
                                            sum_formula_str = f"=SUM({total_col_letter}2:{total_col_letter}{num_data_rows + 1})"
                                            sum_cell = worksheet[f"{total_col_letter}{total_formula_row_num}"]
                                            sum_cell.value = sum_formula_str # Assign formula string
                                            sum_cell.number_format = '#,##0.00 €' # Apply number format
                                            # Optional: Apply bold font to summary rows
                                            # label_col_letter = get_column_letter(export_columns.index(label_col) + 1)
                                            # min_req_row_num = total_formula_row_num + 1
                                            # try:
                                            #    worksheet[f"{label_col_letter}{total_formula_row_num}"].font = Font(bold=True)
                                            #    worksheet[f"{total_col_letter}{total_formula_row_num}"].font = Font(bold=True)
                                            #    worksheet[f"{label_col_letter}{min_req_row_num}"].font = Font(bold=True)
                                            #    worksheet[f"{total_col_letter}{min_req_row_num}"].font = Font(bold=True)
                                            # except NameError: pass # Font not imported
                                            # except Exception as e_font: logging.warning(f"Could not apply font: {e_font}")


                                        sheets_created_count_export += 1
                                        logging.info(f"Export: Successfully processed sheet '{sanitized_name}' with formulas.")

                                    except Exception as write_error:
                                        st.error(f"❌ Erreur lors de l'écriture ou de l'application des formules pour {supplier_export} ({sanitized_name}): {write_error}")
                                        logging.exception(f"Export: Error writing sheet/formulas for {sanitized_name}:") # Log full traceback

                                else: # df_supplier_export_filtered was empty
                                     logging.info(f"Export: No items to order found for supplier '{supplier_export}' in the results, skipping sheet creation.")
                            # End of loop for suppliers
                        else: # Formula not ready
                             st.error("❌ Export annulé car les colonnes nécessaires pour les formules n'ont pas pu être identifiées correctement dans les résultats.")
                except Exception as e_writer:
                    st.error(f"❌ Erreur majeure lors de la création du fichier Excel : {e_writer}")
                    logging.exception("Export: Error during ExcelWriter context:")

                # --- Download Button ---
                if sheets_created_count_export > 0:
                    output_export.seek(0) # Go to the start of the BytesIO buffer
                    # Create filename dynamically
                    export_suppliers_str = "multiples" if len(suppliers_displayed) > 1 else sanitize_sheet_name(suppliers_displayed[0])
                    export_timestamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M')
                    export_filename = f"commande_{export_suppliers_str}_{export_timestamp}.xlsx"

                    st.download_button(
                        label=f"📥 Télécharger Commandes ({sheets_created_count_export} Onglet{'s' if sheets_created_count_export > 1 else ''})",
                        data=output_export, # Use the buffer where Excel data was written
                        file_name=export_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_button"
                        )
                    logging.info(f"Export: Download button created for {sheets_created_count_export} sheets.")
                elif formula_ready: # Formulas were intended but no data rows existed in results
                     st.info("ℹ️ Aucune quantité à commander n'a été trouvée dans les résultats pour l'exportation.")
                     logging.info("Export: No sheets created (no qty > 0 in results).")
                # else: Error occurred before getting here or formula_ready was False

            else: # df_export_all was empty
                st.info("ℹ️ Aucune quantité > 0 trouvée dans les résultats, aucun fichier d'export généré.")
                logging.info("Export: df_export_all was empty, skipping export.")

        else: # No calculation results in session state
            st.info("Veuillez cliquer sur 'Calculer les Quantités à Commander' pour voir les résultats.")


# --- App footer/initial message ---
elif not uploaded_file:
    st.info("👋 Bienvenue ! Chargez votre fichier Excel pour commencer.")

# --- Final catch-all (optional, use with caution) ---
# except Exception as e_global:
#     st.error(f"❌ Une erreur globale et imprévue est survenue dans l'application: {e_global}")
#     logging.exception("Unhandled exception occurred in the main app flow:")
