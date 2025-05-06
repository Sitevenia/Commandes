import streamlit as st
import pandas as pd
import numpy as np
import io
import logging
import re
import openpyxl # Utilisé indirectement par pd.ExcelWriter(engine='openpyxl')
from openpyxl.utils import get_column_letter # Correction import
import calendar

# --- Logging Configuration ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- Helper Functions ---

def safe_read_excel(uploaded_file, sheet_name, **kwargs):
    """ Safely reads an Excel sheet, returning None if sheet not found or error occurs. """
    try:
        # Ensure BytesIO stream is reset if it's already been read
        if isinstance(uploaded_file, io.BytesIO): uploaded_file.seek(0)
        
        # Determine engine based on file extension from the name attribute
        file_name = getattr(uploaded_file, 'name', '')
        engine = 'openpyxl' if file_name.lower().endswith('.xlsx') else None # xlrd for .xls might be needed if pandas default fails

        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, engine=engine, **kwargs)
        
        # Check if DataFrame is truly empty (no columns, no rows)
        # pandas.read_excel can return a DataFrame with columns but no rows if sheet is empty but has headers
        if df.empty and len(df.columns) == 0:
             logging.warning(f"Sheet '{sheet_name}' was read but appears completely empty (no columns, no data).")
             st.warning(f"⚠️ L'onglet '{sheet_name}' semble complètement vide (pas de colonnes, pas de données).")
             return None
        # If it has columns but no data rows, it's "empty" in terms of data content.
        if df.empty and len(df.columns) > 0:
            logging.info(f"Sheet '{sheet_name}' read with columns but no data rows.")
            # This might be acceptable, so we don't return None here unless it's an issue downstream.

        return df
    except ValueError as e:
        if f"Worksheet named '{sheet_name}' not found" in str(e) or f"'{sheet_name}' not found" in str(e):
             logging.warning(f"Sheet '{sheet_name}' not found in the Excel file.")
             st.warning(f"⚠️ Onglet '{sheet_name}' non trouvé dans le fichier Excel.")
        else:
             logging.error(f"ValueError reading sheet '{sheet_name}': {e}")
             st.error(f"❌ Erreur de valeur lors de la lecture de l'onglet '{sheet_name}': {e}.")
        return None
    except FileNotFoundError: # Should not happen with BytesIO but good practice
        logging.error(f"FileNotFoundError (unexpected with BytesIO) reading sheet '{sheet_name}'.")
        st.error(f"❌ Fichier non trouvé (erreur interne) lors de la lecture de l'onglet '{sheet_name}'.")
        return None
    except Exception as e:
        if "zip file" in str(e).lower(): # Often indicates a corrupted .xlsx file
             logging.error(f"Error reading sheet '{sheet_name}': Bad zip file (corrupted .xlsx) - {e}")
             st.error(f"❌ Erreur lors de la lecture de l'onglet '{sheet_name}': Fichier .xlsx potentiellement corrompu (erreur zip).")
        else:
            logging.error(f"Unexpected error reading sheet '{sheet_name}': {type(e).__name__} - {e}")
            st.error(f"❌ Erreur inattendue ('{type(e).__name__}') lors de la lecture de l'onglet '{sheet_name}': {e}.")
        return None

def calculer_quantite_a_commander(df, semaine_columns, montant_minimum_input, duree_semaines):
    """ Calcule la quantité à commander. """
    try:
        if not isinstance(df, pd.DataFrame) or df.empty:
            st.info("Aucune donnée fournie pour le calcul des quantités.")
            return None
            
        required_cols = ["Stock", "Conditionnement", "Tarif d'achat"] + semaine_columns
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            st.error(f"Colonnes manquantes pour le calcul : {', '.join(missing_cols)}")
            return None
            
        if not semaine_columns:
            st.error("Aucune colonne 'semaine' n'a été identifiée pour le calcul des ventes.")
            return None

        df_calc = df.copy()
        # Ensure essential calculation columns are numeric, coercing errors and filling NaNs
        for col in required_cols: # Includes semaine_columns
            df_calc[col] = pd.to_numeric(df_calc[col], errors='coerce').replace([np.inf, -np.inf], np.nan).fillna(0)

        num_semaines_totales = len(semaine_columns)
        ventes_N1 = df_calc[semaine_columns].sum(axis=1)

        if num_semaines_totales >= 64: # Need at least 52+12 weeks for N-1 same period
            v12N1 = df_calc[semaine_columns[-64:-52]].sum(axis=1) # Semaines S-64 à S-53 (12 semaines)
            v12N1s = df_calc[semaine_columns[-52:-40]].sum(axis=1) # Semaines S-52 à S-41 (12 semaines, N-1 "identiques")
            avg12N1 = v12N1 / 12
            avg12N1s = v12N1s / 12
        else:
            v12N1 = pd.Series(0.0, index=df_calc.index)
            v12N1s = pd.Series(0.0, index=df_calc.index)
            avg12N1 = pd.Series(0.0, index=df_calc.index) # Ensure Series for vectorized ops
            avg12N1s = pd.Series(0.0, index=df_calc.index)

        nb_semaines_recentes = min(num_semaines_totales, 12)
        if nb_semaines_recentes > 0:
            v12last = df_calc[semaine_columns[-nb_semaines_recentes:]].sum(axis=1)
            avg12last = v12last / nb_semaines_recentes
        else:
            v12last = pd.Series(0.0, index=df_calc.index)
            avg12last = pd.Series(0.0, index=df_calc.index) # Ensure Series

        # Weighted quantity calculation
        qpond = (0.5 * avg12last + 0.2 * avg12N1 + 0.3 * avg12N1s)
        qnec = qpond * duree_semaines # Necessary quantity before considering stock
        
        # Quantity to order before packaging constraint
        qcomm_series = (qnec - df_calc["Stock"]).apply(lambda x: max(0, x))
        
        cond = df_calc["Conditionnement"]
        stock = df_calc["Stock"]
        tarif = df_calc["Tarif d'achat"]
        
        qcomm = qcomm_series.tolist() # Work with a list for easier item-wise modification

        # Apply packaging constraint
        for i in range(len(qcomm)):
            c = cond.iloc[i]
            q = qcomm[i]
            if q > 0 and c > 0:
                qcomm[i] = int(np.ceil(q / c) * c)
            elif q > 0 and c <= 0: # Cannot order if conditionnement is invalid
                st.warning(f"Article {df_calc.index[i]} (Ref: {df_calc.get('Référence Article', ['N/A'])[i]}) Qté {q:.2f} ignorée car conditionnement est {c}.")
                qcomm[i] = 0 
            else: # q <= 0
                qcomm[i] = 0
        
        # Min order qty if stock is low and recent sales exist
        if nb_semaines_recentes > 0:
            for i in range(len(qcomm)):
                c = cond.iloc[i]
                # Count of recent weeks with sales > 0 for this item
                vr_count = (df_calc[semaine_columns[-nb_semaines_recentes:]].iloc[i] > 0).sum()
                if vr_count >= 2 and stock.iloc[i] <= 1 and c > 0:
                    qcomm[i] = max(qcomm[i], c) # Order at least one packaging unit

        # Zero out order if total N-1 sales and recent sales are very low
        for i in range(len(qcomm)):
            vt_n1_item = ventes_N1.iloc[i]
            vr_sum_item = v12last.iloc[i] # Sum of last 12 weeks sales for this item
            if vt_n1_item < 6 and vr_sum_item < 2:
                qcomm[i] = 0

        # --- Montant Minimum Adjustment ---
        qcomm_df_temp = pd.Series(qcomm, index=df_calc.index) # For easier calculation with tarif
        mt_avant_ajustement = (qcomm_df_temp * tarif).sum()

        if montant_minimum_input > 0 and mt_avant_ajustement < montant_minimum_input:
            mt_actuel = mt_avant_ajustement
            
            # Indices of items that *could* be incremented:
            # - Positive conditioning
            # - Positive price
            # - (Optionally, prioritize those already in qcomm > 0, or allow adding new ones)
            # For this fix, we stick to incrementing items already in qcomm > 0
            
            # Create a list of (original_index, current_qcomm, conditionnement, tarif)
            # for items that are already in the order and can be incremented
            eligible_for_increment = []
            for i in range(len(qcomm)):
                if qcomm[i] > 0 and cond.iloc[i] > 0 and tarif.iloc[i] > 0:
                    eligible_for_increment.append(i) # Store original index

            if not eligible_for_increment:
                if mt_actuel < montant_minimum_input:
                    st.warning(
                        f"Impossible d'atteindre le montant minimum de {montant_minimum_input:,.2f}€. "
                        f"Montant actuel: {mt_actuel:,.2f}€. "
                        "Aucun article commandé avec conditionnement et tarif valides pour incrémentation."
                    )
            else:
                idx_ptr_eligible = 0 # Pointer for the 'eligible_for_increment' list
                # Max iterations to prevent infinite loops (e.g. if all prices are 0)
                # Iterate at most (e.g.) 20 times over all eligible items
                max_iter_loop = len(eligible_for_increment) * 20 + 1 
                iters = 0

                while mt_actuel < montant_minimum_input and iters < max_iter_loop:
                    iters += 1
                    
                    # Get the actual DataFrame index of the item to increment
                    original_df_idx = eligible_for_increment[idx_ptr_eligible]
                    
                    c_item = cond.iloc[original_df_idx] # Known to be > 0
                    p_item = tarif.iloc[original_df_idx] # Known to be > 0
                    
                    # Increment quantity for this item by its packaging unit
                    qcomm[original_df_idx] += c_item
                    mt_actuel += c_item * p_item # Update current total amount
                    
                    # Move to the next eligible item, cycling through the list
                    idx_ptr_eligible = (idx_ptr_eligible + 1) % len(eligible_for_increment)
                
                if iters >= max_iter_loop and mt_actuel < montant_minimum_input:
                    st.error(
                        f"Ajustement du montant minimum : Nombre maximum d'itérations ({max_iter_loop}) atteint. "
                        f"Montant actuel: {mt_actuel:,.2f}€ / Requis: {montant_minimum_input:,.2f}€. "
                        "Vérifiez les tarifs et conditionnements."
                    )
        
        # Recalculate final total amount after potential adjustments
        qcomm_final_series = pd.Series(qcomm, index=df_calc.index)
        mt_final = (qcomm_final_series * tarif).sum()
        
        return (qcomm, ventes_N1, v12N1, v12last, mt_final)

    except KeyError as e:
        st.error(f"Erreur de clé (colonne manquante probable) lors du calcul des quantités : '{e}'. Vérifiez les noms des colonnes dans le fichier.")
        logging.exception(f"KeyError in calculer_quantite_a_commander: {e}")
        return None
    except Exception as e:
        st.error(f"Erreur inattendue lors du calcul des quantités : {type(e).__name__} - {e}")
        logging.exception("Exception in calculer_quantite_a_commander:")
        return None


def calculer_rotation_stock(df, semaine_columns, periode_semaines):
    """ Calcule les métriques de rotation de stock. """
    try:
        if not isinstance(df, pd.DataFrame) or df.empty:
            st.info("Aucune donnée fournie pour l'analyse de rotation.")
            return pd.DataFrame() # Return empty DataFrame for consistency

        required_cols = ["Stock", "Tarif d'achat"] # Semaine_columns checked separately
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            st.error(f"Colonnes manquantes pour l'analyse de rotation : {', '.join(missing_cols)}")
            return None

        df_rotation = df.copy()

        # Determine analysis period for sales
        if periode_semaines and periode_semaines > 0 and len(semaine_columns) >= periode_semaines:
            semaines_analyse = semaine_columns[-periode_semaines:]
            nb_semaines_analyse = periode_semaines
        elif periode_semaines and periode_semaines > 0: # Not enough history for requested period
            semaines_analyse = semaine_columns
            nb_semaines_analyse = len(semaine_columns)
            st.caption(f"Période d'analyse ajustée à {nb_semaines_analyse} semaines (données disponibles).")
        else: # periode_semaines is 0 or invalid, use all available
            semaines_analyse = semaine_columns
            nb_semaines_analyse = len(semaine_columns)
        
        if not semaines_analyse:
            st.warning("Aucune colonne de vente disponible pour l'analyse de rotation.")
            # Return df_rotation with original columns, possibly adding empty metric columns
            # For simplicity, we might return it as is or add empty metric columns.
            # Let's add empty columns for consistency if processing continues.
            metric_cols = ["Unités Vendues (Période)", "Ventes Moy Hebdo (Période)", "Ventes Moy Mensuel (Période)",
                           "Semaines Stock (WoS)", "Rotation Unités (Proxy)", "COGS (Période)", 
                           "Valeur Stock Actuel (€)", "Rotation Valeur (Proxy)"]
            for m_col in metric_cols: df_rotation[m_col] = 0.0
            return df_rotation

        # Ensure sales columns are numeric
        for col in semaines_analyse:
            df_rotation[col] = pd.to_numeric(df_rotation[col], errors='coerce').fillna(0)

        df_rotation["Unités Vendues (Période)"] = df_rotation[semaines_analyse].sum(axis=1)
        
        if nb_semaines_analyse > 0:
            df_rotation["Ventes Moy Hebdo (Période)"] = df_rotation["Unités Vendues (Période)"] / nb_semaines_analyse
        else:
            df_rotation["Ventes Moy Hebdo (Période)"] = 0.0
            
        avg_weeks_per_month = 52 / 12
        df_rotation["Ventes Moy Mensuel (Période)"] = df_rotation["Ventes Moy Hebdo (Période)"] * avg_weeks_per_month
        
        df_rotation["Stock"] = pd.to_numeric(df_rotation["Stock"], errors='coerce').fillna(0)
        df_rotation["Tarif d'achat"] = pd.to_numeric(df_rotation["Tarif d'achat"], errors='coerce').fillna(0)
        
        # Weeks of Stock (WoS)
        denom_wos = df_rotation["Ventes Moy Hebdo (Période)"]
        df_rotation["Semaines Stock (WoS)"] = np.divide(
            df_rotation["Stock"], denom_wos, 
            out=np.full_like(df_rotation["Stock"], np.inf, dtype=np.float64), # Fill with inf where denom is 0
            where=denom_wos != 0
        )
        df_rotation.loc[df_rotation["Stock"] <= 0, "Semaines Stock (WoS)"] = 0.0 # If no stock, WoS is 0

        # Rotation Unités (Proxy)
        # Using (Units Sold Period) / Avg Stock. Here, Avg Stock is proxied by Current Stock.
        denom_rot_unit = df_rotation["Stock"] # Proxy for average stock
        df_rotation["Rotation Unités (Proxy)"] = np.divide(
            df_rotation["Unités Vendues (Période)"], denom_rot_unit,
            out=np.full_like(denom_rot_unit, np.inf, dtype=np.float64),
            where=denom_rot_unit != 0
        )
        # If no units sold AND no stock, rotation is undefined, typically 0 in reports
        df_rotation.loc[(df_rotation["Unités Vendues (Période)"] <= 0) & (denom_rot_unit <= 0), "Rotation Unités (Proxy)"] = 0.0
        # If units sold but no stock, rotation is inf (already handled by np.divide)
        # If no units sold but stock exists, rotation is 0
        df_rotation.loc[(df_rotation["Unités Vendues (Période)"] <= 0) & (denom_rot_unit > 0), "Rotation Unités (Proxy)"] = 0.0


        df_rotation["COGS (Période)"] = df_rotation["Unités Vendues (Période)"] * df_rotation["Tarif d'achat"]
        df_rotation["Valeur Stock Actuel (€)"] = df_rotation["Stock"] * df_rotation["Tarif d'achat"]
        
        # Rotation Valeur (Proxy)
        # Using (COGS Period) / Avg Stock Value. Avg Stock Value proxied by Current Stock Value.
        denom_rot_val = df_rotation["Valeur Stock Actuel (€)"] # Proxy for average stock value
        df_rotation["Rotation Valeur (Proxy)"] = np.divide(
            df_rotation["COGS (Période)"], denom_rot_val,
            out=np.full_like(denom_rot_val, np.inf, dtype=np.float64),
            where=denom_rot_val != 0
        )
        # If COGS is 0 AND stock value is 0, rotation is undefined, typically 0
        df_rotation.loc[(df_rotation["COGS (Période)"] <= 0) & (denom_rot_val <= 0), "Rotation Valeur (Proxy)"] = 0.0
        # If COGS > 0 but stock value is 0, rotation is inf
        # If COGS is 0 but stock value > 0, rotation is 0
        df_rotation.loc[(df_rotation["COGS (Période)"] <= 0) & (denom_rot_val > 0), "Rotation Valeur (Proxy)"] = 0.0

        # df_rotation["Rotation Unités (Proxy)"].fillna(0, inplace=True) # Covered by specific loc conditions
        # df_rotation["Rotation Valeur (Proxy)"].fillna(0, inplace=True) # Covered

        return df_rotation

    except KeyError as e:
        st.error(f"Erreur de clé (colonne manquante probable) lors du calcul de la rotation : '{e}'.")
        logging.exception(f"KeyError in calculer_rotation_stock: {e}")
        return None
    except Exception as e:
        st.error(f"Erreur inattendue lors du calcul de la rotation : {type(e).__name__} - {e}")
        logging.exception("Error in calculer_rotation_stock:")
        return None

def approx_weeks_to_months(week_columns_52):
    """Approximates month mapping for 52 consecutive week columns."""
    month_map = {}
    if not week_columns_52 or len(week_columns_52) != 52:
        logging.warning(f"approx_weeks_to_months expects 52 columns, got {len(week_columns_52)}. Returning empty map.")
        return month_map

    weeks_per_month_approx = 52 / 12.0 # Use float for precision
    current_week_index = 0
    
    for i in range(1, 13): # For months 1 to 12
        month_name = calendar.month_name[i]
        # Determine number of full weeks for this month (4 or 5)
        # A common approximation: 4-4-5 pattern for weeks in months per quarter
        # Or, more simply, distribute based on weeks_per_month_approx
        
        start_idx_exact = (i-1) * weeks_per_month_approx
        end_idx_exact = i * weeks_per_month_approx
        
        # Round to nearest integer for start, and ensure end_idx is at least start_idx + number of weeks
        # A simpler way is to accumulate week counts
        start_idx_for_month = int(round(start_idx_exact))
        end_idx_for_month = int(round(end_idx_exact))

        # Ensure indices are within bounds of the 52 weeks
        start_idx_for_month = max(0, start_idx_for_month)
        end_idx_for_month = min(len(week_columns_52), end_idx_for_month)
        
        # Ensure end_idx is at least start_idx
        if end_idx_for_month < start_idx_for_month:
             # This can happen due to rounding if weeks_per_month_approx is small for a particular month's slice
             # Default to at least one week if it's meant to be a segment
             # However, the provided code's original logic works well:
             # start_idx = int(round((i-1) * weeks_per_month_approx))
             # end_idx = int(round(i * weeks_per_month_approx))

            # Using the original logic as it's simpler and likely sufficient for approximation
            start_idx = int(round((i-1) * weeks_per_month_approx))
            end_idx = int(round(i * weeks_per_month_approx))
            month_cols = week_columns_52[start_idx : min(end_idx, 52)] # Ensure end_idx doesn't exceed 52
            month_map[month_name] = month_cols

    logging.info(f"Approximated month-to-week map created. Example January: {month_map.get('January', [])}")
    return month_map


def calculer_forecast_simulation_v3(df, all_semaine_columns, selected_months, sim_type, progression_pct=0, objectif_montant=0):
    """ Performs forecast simulation for SELECTED MONTHS based on corresponding N-1 data. """
    try:
        if not isinstance(df, pd.DataFrame) or df.empty:
            st.warning("Aucune donnée fournie pour la simulation de forecast.")
            return None, 0.0

        if not all_semaine_columns or len(all_semaine_columns) < 52:
            st.error("Données historiques insuffisantes (< 52 semaines identifiées) pour une simulation basée sur N-1.")
            return None, 0.0

        if not selected_months:
            st.warning("Veuillez sélectionner au moins un mois pour la simulation.")
            return None, 0.0

        required_cols = ["Référence Article", "Désignation Article", "Conditionnement", "Tarif d'achat", "Fournisseur"] # Added Fournisseur
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            st.error(f"Colonnes manquantes pour la simulation : {', '.join(missing_cols)}")
            return None, 0.0

        # Identify N-1 columns by parsing year and week from column names
        years_in_cols = set()
        parsed_week_cols = [] # List of (year, week_num, original_col_name)
        for col_name in all_semaine_columns:
            if isinstance(col_name, str):
                # Regex to capture YYYYWW or YYYYSW (S for semaine)
                match = re.match(r"(\d{4})S?(\d{1,2})", col_name, re.IGNORECASE)
                if match:
                    year = int(match.group(1))
                    week = int(match.group(2))
                    if 1 <= week <= 53: # Valid week number
                        years_in_cols.add(year)
                        parsed_week_cols.append({'year': year, 'week': week, 'col': col_name, 'sort_key': year * 100 + week})
        
        if not years_in_cols:
            st.error("Impossible de déterminer les années à partir des noms de colonnes semaines. Format attendu 'YYYYWW' ou 'YYYYSwW'.")
            return None, 0.0

        # Sort parsed_week_cols by year then week to ensure correct order
        parsed_week_cols.sort(key=lambda x: x['sort_key'])
        
        # Assuming current year (N) is the max year found in data, N-1 is that year minus 1.
        # This assumes the data contains the "current" period or recent data.
        year_n = max(years_in_cols) if years_in_cols else 0
        year_n_minus_1 = year_n - 1
        
        st.caption(f"Simulation basée sur N-1 (Année N détectée: {year_n}, Année N-1 utilisée: {year_n_minus_1})")

        # Filter for N-1 week columns from the *sorted* list
        n1_week_cols_data = [item for item in parsed_week_cols if item['year'] == year_n_minus_1]
        
        if len(n1_week_cols_data) < 52: # Need full 52 weeks for N-1
            st.error(f"Données N-1 ({year_n_minus_1}) insuffisantes. {len(n1_week_cols_data)} semaines trouvées, 52 requises.")
            return None, 0.0
        
        # Take the first 52 weeks of N-1 for consistent monthly mapping
        n1_week_cols_for_mapping = [item['col'] for item in n1_week_cols_data[:52]]

        df_sim = df[required_cols].copy() # Select only necessary columns for simulation output base
        df_sim["Tarif d'achat"] = pd.to_numeric(df_sim["Tarif d'achat"], errors='coerce').fillna(0)
        df_sim["Conditionnement"] = pd.to_numeric(df_sim["Conditionnement"], errors='coerce').fillna(1)
        df_sim["Conditionnement"] = df_sim["Conditionnement"].apply(lambda x: 1 if x <= 0 else int(x))


        # Ensure all N-1 columns for mapping are actually in the input DataFrame's sales data part
        missing_n1_in_df = [col for col in n1_week_cols_for_mapping if col not in df.columns]
        if missing_n1_in_df:
            st.error(f"Erreur interne: Colonnes N-1 mappées ({', '.join(missing_n1_in_df)}) non trouvées dans les données de vente du DataFrame.")
            return None, 0.0
            
        df_n1_sales_data = df[n1_week_cols_for_mapping].copy() # Use only the 52 N-1 weeks for sales sum
        for col in n1_week_cols_for_mapping: # Ensure numeric
            df_n1_sales_data[col] = pd.to_numeric(df_n1_sales_data[col], errors='coerce').fillna(0)

        # --- Monthly Seasonality based on N-1 (first 52 weeks) ---
        month_col_map_n1 = approx_weeks_to_months(n1_week_cols_for_mapping) # Uses the 52 N-1 week cols
        
        total_n1_sales_selected_months_series = pd.Series(0.0, index=df_sim.index)
        monthly_sales_n1_for_selected_months = {} # Store N-1 sales Series for each selected month

        for month_name in selected_months:
            if month_name in month_col_map_n1 and month_col_map_n1[month_name]:
                # Get actual N-1 sales columns corresponding to this month's approximation
                actual_cols_for_month_n1 = [col for col in month_col_map_n1[month_name] if col in df_n1_sales_data.columns]
                if actual_cols_for_month_n1:
                    sales_this_month_n1 = df_n1_sales_data[actual_cols_for_month_n1].sum(axis=1)
                    monthly_sales_n1_for_selected_months[month_name] = sales_this_month_n1
                    total_n1_sales_selected_months_series += sales_this_month_n1
                    df_sim[f"Ventes N-1 {month_name}"] = sales_this_month_n1
                else: # No N-1 columns mapped for this month (should not happen if map is good)
                    monthly_sales_n1_for_selected_months[month_name] = pd.Series(0.0, index=df_sim.index)
                    df_sim[f"Ventes N-1 {month_name}"] = 0.0
            else: # Month not in map or map has no columns for it
                monthly_sales_n1_for_selected_months[month_name] = pd.Series(0.0, index=df_sim.index)
                df_sim[f"Ventes N-1 {month_name}"] = 0.0
        
        df_sim["Vts N-1 Tot (Mois Sel.)"] = total_n1_sales_selected_months_series

        # Calculate seasonality factor for each selected month relative to total N-1 sales of *selected months*
        period_seasonality_factors = {}
        # Use .replace(0, np.nan) for safe division, then .fillna(0) if total is 0
        safe_total_n1_for_selected_months = total_n1_sales_selected_months_series.copy()
        # If total N-1 sales for the selected period is 0 for an item, seasonality is undefined or could be equal.
        # To avoid division by zero, replace 0s with NaN, then division results in NaN, then fill NaN with 0.
        # If total is 0, but a month had sales (impossible), this logic still holds.
        # If all selected months have 0 sales, all seasonality factors will be 0.
        
        # For items where total_n1_sales_selected_months_series is 0, their seasonality factors will be 0.
        # If we want to distribute equally in that case, more logic is needed here.
        # Current: if total N-1 for selected period is 0, then forecast based on seasonality will also be 0.

        for month_name in selected_months:
            if month_name in monthly_sales_n1_for_selected_months:
                # Calculate seasonality: (Month's N-1 Sales) / (Total N-1 Sales for ALL Selected Months)
                # This gives the proportion of sales for that month within the selected period.
                # np.divide handles 0 in denominator by outputting np.nan or inf, which we then manage.
                month_sales_n1 = monthly_sales_n1_for_selected_months[month_name]
                factor = np.divide(month_sales_n1, safe_total_n1_for_selected_months, 
                                   out=np.zeros_like(month_sales_n1, dtype=float), # Output 0 where denom is 0
                                   where=safe_total_n1_for_selected_months!=0)
                period_seasonality_factors[month_name] = pd.Series(factor, index=df_sim.index).fillna(0)
            else:
                period_seasonality_factors[month_name] = pd.Series(0.0, index=df_sim.index)


        # --- Calculate Forecasted Quantities ---
        base_monthly_forecast_qty_map = {} # Store base forecasted QTY (before packaging)

        if sim_type == 'Simple Progression':
            prog_factor = 1 + (progression_pct / 100.0)
            # Total forecasted quantity for the entire selected period
            total_forecast_qty_for_selected_period = total_n1_sales_selected_months_series * prog_factor
            
            for month_name in selected_months:
                # Distribute total forecast based on N-1 seasonality of that month within the selected period
                seasonality_for_month = period_seasonality_factors.get(month_name, pd.Series(0.0, index=df_sim.index))
                base_monthly_forecast_qty_map[month_name] = total_forecast_qty_for_selected_period * seasonality_for_month
        
        elif sim_type == 'Objectif Montant':
            if objectif_montant <= 0:
                st.error("Pour 'Objectif Montant', l'objectif doit être supérieur à 0.")
                return None, 0.0

            # Sum of N-1 sales (units) across all items for the selected period
            total_n1_sales_units_all_items = total_n1_sales_selected_months_series.sum()

            if total_n1_sales_units_all_items <= 0: # No N-1 sales at all for any item in selected period
                st.warning(
                    "Les ventes N-1 pour les mois sélectionnés sont nulles pour tous les articles. "
                    "Tentative de répartition égale du montant objectif sur les mois et les articles (si tarif > 0)."
                )
                num_selected_months = len(selected_months)
                if num_selected_months == 0: return None, 0.0 # Should have been caught earlier
                
                # Equally distribute the target amount per month
                target_amount_per_month_for_all_items = objectif_montant / num_selected_months
                
                for month_name in selected_months:
                    # For each item, its share of this monthly amount is proportional to its price (kind of inverse logic here)
                    # Simpler: assume each item gets an equal share of units IF price was uniform.
                    # Better: distribute target_amount_per_month_for_all_items across items based on their price.
                    # Or, if we must get qty: Qty = Target_Amount_for_Item / Price_Item
                    # This means items with lower price get higher qty for same target amount portion.
                    # This allocation is tricky without a clear N-1 unit or value basis.
                    # Simplest for "no N-1 sales": allocate target_amount_per_month to items with price > 0.
                    # Here, we'll try to assign the monthly target amount *per item* based on its price,
                    # effectively making each item aim for target_amount_per_month_for_all_items / num_items_with_price.
                    # The provided code calculates qty = (target_amount_for_month * seasonality) / price.
                    # If N-1 sales are zero, seasonality is zero, so qty becomes zero. This needs adjustment.

                    # If total N-1 sales are zero, seasonality factors are all zero.
                    # We need an alternative way to distribute the objectif_montant.
                    # Let's distribute the objectif_montant for this month (total_obj / num_months)
                    # equally among items that have a price.
                    
                    num_items_with_price = (df_sim["Tarif d'achat"] > 0).sum()
                    if num_items_with_price == 0:
                        base_monthly_forecast_qty_map[month_name] = pd.Series(0.0, index=df_sim.index)
                    else:
                        target_amount_per_item_this_month = target_amount_per_month_for_all_items / num_items_with_price
                        base_monthly_forecast_qty_map[month_name] = np.divide(
                            target_amount_per_item_this_month, df_sim["Tarif d'achat"],
                            out=np.zeros_like(df_sim["Tarif d'achat"], dtype=float),
                            where=df_sim["Tarif d'achat"] != 0
                        )
            else: # Normal case: N-1 sales exist for the period for at least some items
                for month_name in selected_months:
                    seasonality_for_month = period_seasonality_factors.get(month_name, pd.Series(0.0, index=df_sim.index))
                    # Target amount for this specific month for each item (vectorized)
                    target_amount_for_this_month_per_item = objectif_montant * seasonality_for_month
                    
                    base_monthly_forecast_qty_map[month_name] = np.divide(
                        target_amount_for_this_month_per_item, df_sim["Tarif d'achat"],
                        out=np.zeros_like(df_sim["Tarif d'achat"], dtype=float), # Qty is 0 if price is 0
                        where=df_sim["Tarif d'achat"] != 0
                    )
        else:
            st.error(f"Type de simulation non reconnu : '{sim_type}'.")
            return None, 0.0

        # --- Adjust by Packaging (Conditionnement) and Calculate Final Amounts ---
        total_adjusted_qty_all_months = pd.Series(0.0, index=df_sim.index)
        total_final_amount_all_months = pd.Series(0.0, index=df_sim.index)

        for month_name in selected_months:
            forecast_qty_col_name = f"Qté Prév. {month_name}"
            forecast_amount_col_name = f"Montant Prév. {month_name} (€)"
            
            if month_name in base_monthly_forecast_qty_map:
                base_q_series = pd.to_numeric(base_monthly_forecast_qty_map[month_name], errors='coerce').fillna(0)
                cond_series = df_sim["Conditionnement"] # Already int, >0
                
                # Adjusted quantity = ceil(base_quantity / conditionnement) * conditionnement
                adjusted_qty_series = (
                    np.ceil(
                        np.divide(base_q_series, cond_series, 
                                  out=np.zeros_like(base_q_series, dtype=float), 
                                  where=cond_series != 0) # Should always be !=0 due to prior processing
                    ) * cond_series
                ).fillna(0).astype(int)
                
                df_sim[forecast_qty_col_name] = adjusted_qty_series
                df_sim[forecast_amount_col_name] = adjusted_qty_series * df_sim["Tarif d'achat"]
                
                total_adjusted_qty_all_months += adjusted_qty_series
                total_final_amount_all_months += df_sim[forecast_amount_col_name]
            else: # Should not happen if selected_months drives the loop
                df_sim[forecast_qty_col_name] = 0
                df_sim[forecast_amount_col_name] = 0.0
        
        df_sim["Qté Totale Prév. (Mois Sel.)"] = total_adjusted_qty_all_months
        df_sim["Montant Total Prév. (€) (Mois Sel.)"] = total_final_amount_all_months

        # --- Prepare Final Output DataFrame ---
        # Order of columns for display
        id_cols_display = ["Fournisseur", "Référence Article", "Désignation Article", "Conditionnement", "Tarif d'achat"]
        
        # Columns for N-1 sales for selected months
        n1_sales_cols_display = sorted([f"Ventes N-1 {m}" for m in selected_months if f"Ventes N-1 {m}" in df_sim.columns])
        
        # Columns for forecasted quantities for selected months
        qty_forecast_cols_display = sorted([f"Qté Prév. {m}" for m in selected_months if f"Qté Prév. {m}" in df_sim.columns])
        
        # Columns for forecasted amounts for selected months
        amt_forecast_cols_display = sorted([f"Montant Prév. {m} (€)" for m in selected_months if f"Montant Prév. {m} (€)" in df_sim.columns])
        
        # Columns for totals over the selected period
        total_cols_display = [
            "Vts N-1 Tot (Mois Sel.)", 
            "Qté Totale Prév. (Mois Sel.)", 
            "Montant Total Prév. (€) (Mois Sel.)"
        ]
        # Rename for brevity in UI if desired (original code did this)
        df_sim.rename(columns={
            "Vts N-1 Tot (Mois Sel.)": "Vts N-1 Tot (Mois Sel.)", # No change, was already fine
            "Qté Totale Prév. (Mois Sel.)": "Qté Tot Prév (Mois Sel.)",
            "Montant Total Prév. (€) (Mois Sel.)": "Mnt Tot Prév (€) (Mois Sel.)"
        }, inplace=True)
        # Update total_cols_display if names changed
        total_cols_display = [ # Re-fetch with new names
            "Vts N-1 Tot (Mois Sel.)",
            "Qté Tot Prév (Mois Sel.)",
            "Mnt Tot Prév (€) (Mois Sel.)"
        ]

        final_ordered_cols = id_cols_display + total_cols_display + n1_sales_cols_display + qty_forecast_cols_display + amt_forecast_cols_display
        
        # Ensure all columns in final_ordered_cols actually exist in df_sim to prevent KeyErrors
        final_ordered_cols_existing = [col for col in final_ordered_cols if col in df_sim.columns]

        grand_total_forecast_amount = total_final_amount_all_months.sum()
        
        return df_sim[final_ordered_cols_existing], grand_total_forecast_amount

    except KeyError as e:
        st.error(f"Erreur de clé (colonne manquante probable) lors de la simulation forecast : '{e}'.")
        logging.exception(f"KeyError in calculer_forecast_simulation_v3: {e}")
        return None, 0.0
    except Exception as e:
        st.error(f"Erreur inattendue lors de la simulation forecast : {type(e).__name__} - {e}")
        logging.exception("Error in calculer_forecast_simulation_v3:")
        return None, 0.0


def sanitize_sheet_name(name):
    """ Removes invalid characters for Excel sheet names and truncates to 31 chars. """
    if not isinstance(name, str):
        name = str(name)
    
    # Invalid characters in Excel sheet names: [ ] : * ? / \
    # Excel also doesn't like sheet names starting or ending with an apostrophe if used for quoting.
    # Let's also replace < > | " just to be safe, though they might be less common.
    sanitized = re.sub(r'[\[\]:*?/\\<>|"]', '_', name)
    
    # Sheet names cannot start or end with a single quote '
    if sanitized.startswith("'"):
        sanitized = "_" + sanitized[1:]
    if sanitized.endswith("'"):
        sanitized = sanitized[:-1] + "_"
        
    # Truncate to maximum 31 characters
    return sanitized[:31]

# --- NEW: Function to render supplier checkboxes ---
def render_supplier_checkboxes(tab_key_prefix, all_suppliers, default_select_all=False):
    """
    Renders supplier checkboxes within an expander for a specific tab.
    Manages state using keys based on tab_key_prefix.
    Returns the list of selected suppliers for this tab.
    """
    select_all_key = f"{tab_key_prefix}_select_all"
    # Create a dictionary of supplier_name: session_state_key_for_checkbox
    supplier_cb_keys = {
        supplier: f"{tab_key_prefix}_cb_{sanitize_supplier_key(supplier)}" for supplier in all_suppliers
    }

    # --- State Initialization ---
    # Initialize "Select All" checkbox state for this tab if it doesn't exist
    # This happens on first run or after a full session_state clear for this tab_key_prefix
    if select_all_key not in st.session_state:
        st.session_state[select_all_key] = default_select_all
        # When "Select All" is initialized, also initialize individual checkboxes
        # only if they don't already have a state (e.g., from a previous interaction
        # before a partial state clear that might have missed them).
        for cb_key in supplier_cb_keys.values():
            if cb_key not in st.session_state:
                st.session_state[cb_key] = default_select_all
    else:
        # If "Select All" exists, ensure individual checkboxes also exist.
        # This handles cases where new suppliers might be added to an existing session.
        # Or if state was cleared incompletely.
        # The value will be their current state or default to "Select All"'s current state if new.
        for cb_key in supplier_cb_keys.values():
            if cb_key not in st.session_state:
                 st.session_state[cb_key] = st.session_state[select_all_key]


    # --- Callbacks for Checkbox Interactions ---
    def toggle_all_suppliers_for_tab():
        """Callback for the 'Select All' checkbox of this specific tab."""
        current_select_all_value = st.session_state[select_all_key]
        logging.debug(f"Tab '{tab_key_prefix}': 'Select All' toggled to {current_select_all_value}.")
        for cb_key in supplier_cb_keys.values():
            st.session_state[cb_key] = current_select_all_value

    def check_individual_supplier_for_tab():
        """Callback for individual supplier checkboxes of this specific tab."""
        # Check if all individual checkboxes for this tab are checked
        all_individual_checked = all(
            st.session_state.get(cb_key, False) for cb_key in supplier_cb_keys.values()
        )
        # Update the "Select All" state for this tab if necessary, without triggering its own callback
        if st.session_state[select_all_key] != all_individual_checked:
            st.session_state[select_all_key] = all_individual_checked
            logging.debug(f"Tab '{tab_key_prefix}': 'Select All' auto-updated to {all_individual_checked} due to individual change.")


    # --- Display Widgets ---
    with st.expander("👤 Sélectionner Fournisseurs", expanded=True):
        st.checkbox(
            "Sélectionner / Désélectionner Tout",
            key=select_all_key,
            on_change=toggle_all_suppliers_for_tab, # This callback updates individual checkboxes
            disabled=not bool(all_suppliers) # Disable if no suppliers to select
        )
        st.markdown("---")

        selected_suppliers_in_ui = []
        num_display_cols = 4 # Number of columns for checkbox layout
        checkbox_cols = st.columns(num_display_cols)
        current_col_idx = 0
        
        for supplier_name, cb_key in supplier_cb_keys.items():
            # Use the value directly from session_state for the checkbox's current state
            # The on_change callback for individual boxes will update the "Select All" state
            is_checked = checkbox_cols[current_col_idx].checkbox(
                supplier_name,
                key=cb_key,
                # value parameter is not strictly needed if key reflects state,
                # but explicit for clarity / if default behavior changes.
                # Streamlit checkbox uses st.session_state[key] if value is not provided.
                # value=st.session_state.get(cb_key), # Ensure value reflects actual state
                on_change=check_individual_supplier_for_tab
            )
            if st.session_state.get(cb_key): # Read the state after potential change by user
                selected_suppliers_in_ui.append(supplier_name)
            
            current_col_idx = (current_col_idx + 1) % num_display_cols

    logging.debug(f"Checkboxes rendered for tab '{tab_key_prefix}'. UI selected: {len(selected_suppliers_in_ui)} suppliers.")
    return selected_suppliers_in_ui


def sanitize_supplier_key(supplier_name):
     """Creates a safe key for session state from supplier name."""
     if not isinstance(supplier_name, str):
         supplier_name = str(supplier_name)
     # Replace non-alphanumeric characters (and spaces) with underscores
     s = re.sub(r'\W+', '_', supplier_name)
     # Remove leading/trailing underscores that might result
     s = re.sub(r'^_+|_+$', '', s)
     # Ensure not empty and handle if it becomes all underscores
     s = re.sub(r'_+', '_', s) # Consolidate multiple underscores
     return s if s else "invalid_supplier_key"

# --- Streamlit App Main Logic ---
st.set_page_config(page_title="Forecast & Rotation App", layout="wide")
st.title("📦 Application Prévision Commande & Analyse Rotation")

# --- File Upload ---
# Use a unique key for the uploader if you plan to refer to its state elsewhere,
# though typically its value is processed immediately.
uploaded_file = st.file_uploader(
    "📁 Charger le fichier Excel principal contenant 'Tableau final' et 'Minimum de commande'", 
    type=["xlsx", "xls"], 
    key="main_file_uploader"
)

# --- Initialize Session State for application data and UI controls ---
# Encapsulate default values in a function or dict for clarity
def get_default_session_state():
    return {
        'df_full': None,                # Full data from 'Tableau final'
        'min_order_dict': {},           # Min order amounts per supplier
        'df_initial_filtered': pd.DataFrame(), # Data after initial structural filters
        'all_available_semaine_columns': [],   # All identified sales week columns from data
        'unique_suppliers_list': [],    # Sorted list of all unique suppliers from data

        # Results from Tab 1 (Order Prediction)
        'commande_result_df': None,
        'commande_calculated_total_amount': 0.0,
        'commande_suppliers_calculated_for': [], # Suppliers used for the last calculation

        # Results from Tab 2 (Stock Rotation)
        'rotation_result_df': None,
        'rotation_analysis_period_label': "",
        'rotation_suppliers_calculated_for': [],

        # UI state for Tab 2
        'rotation_threshold_value': 1.0, # Default filter threshold
        'show_all_rotation_data': True,  # Default to show all items

        # Results from Tab 4 (Forecast Simulation)
        'forecast_result_df': None,
        'forecast_grand_total_amount': 0.0,
        'forecast_simulation_params_calculated_for': {}, # Params used for last calculation

        # UI state for Tab 4
        'forecast_selected_months_ui': list(calendar.month_name)[1:], # Default all months
        'forecast_sim_type_radio_index': 0, # Default to 'Simple Progression'
        'forecast_progression_percentage_ui': 5.0,
        'forecast_target_amount_ui': 10000.0,
        
        # Note: Checkbox states (tabX_select_all, tabX_cb_supplier) are managed dynamically
        # by render_supplier_checkboxes and cleared via prefix matching on new file upload.
    }

# Initialize session state keys if they don't exist
for key, default_value in get_default_session_state().items():
    if key not in st.session_state:
        st.session_state[key] = default_value


# --- Data Loading and Initial Processing Block ---
# This block runs if a file is uploaded AND df_full hasn't been populated yet from this file.
# It effectively processes a new file upload.
if uploaded_file and st.session_state.df_full is None:
    logging.info(f"New file uploaded: {uploaded_file.name}. Starting processing...")
    
    # --- Clear previous data and relevant UI states ---
    # Keys for dataframes, results, and specific UI elements tied to data content
    keys_to_reset_on_new_file = [
        'df_full', 'min_order_dict', 'df_initial_filtered', 'all_available_semaine_columns',
        'unique_suppliers_list',
        'commande_result_df', 'commande_calculated_total_amount', 'commande_suppliers_calculated_for',
        'rotation_result_df', 'rotation_analysis_period_label', 'rotation_suppliers_calculated_for',
        'forecast_result_df', 'forecast_grand_total_amount', 'forecast_simulation_params_calculated_for',
        # Add any other data-dependent state keys here
    ]
    # Also clear dynamically created checkbox states (prefixed keys)
    # These prefixes are used in render_supplier_checkboxes
    dynamic_key_prefixes_to_clear = ['tab1_', 'tab2_', 'tab4_'] 
                                     # Add 'tab3_' if it ever gets supplier checkboxes

    for key in keys_to_reset_on_new_file:
        if key in st.session_state:
            del st.session_state[key]
            logging.debug(f"Reset key: {key}")
    
    for prefix in dynamic_key_prefixes_to_clear:
        keys_to_remove = [k for k in st.session_state if k.startswith(prefix)]
        for k_to_remove in keys_to_remove:
            del st.session_state[k_to_remove]
            logging.debug(f"Reset dynamic key: {k_to_remove}")

    # Re-initialize with defaults after clearing
    for key, default_value in get_default_session_state().items():
        st.session_state[key] = default_value
    logging.info("Session state cleared and re-initialized for new file.")

    # --- Read and process the uploaded Excel file ---
    try:
        # Use a BytesIO buffer to handle the uploaded file in memory
        excel_file_buffer = io.BytesIO(uploaded_file.getvalue())
        
        st.info("Lecture de l'onglet 'Tableau final'...")
        df_full_temp = safe_read_excel(excel_file_buffer, sheet_name="Tableau final", header=7)
        
        if df_full_temp is None:
            st.error("❌ Échec de la lecture de l'onglet 'Tableau final'. Vérifiez le nom de l'onglet et le format du fichier.")
            st.stop() # Stop execution if essential data is missing

        # Validate required columns in 'Tableau final'
        required_cols_tableau_final = ["Stock", "Fournisseur", "AF_RefFourniss", "Tarif d'achat", "Conditionnement", "Référence Article", "Désignation Article"]
        missing_cols_tf = [col for col in required_cols_tableau_final if col not in df_full_temp.columns]
        if missing_cols_tf:
            st.error(f"❌ Colonnes manquantes dans 'Tableau final': {', '.join(missing_cols_tf)}. L'application ne peut pas continuer.")
            st.stop()

        # Basic data cleaning and type conversion for essential columns
        df_full_temp["Stock"] = pd.to_numeric(df_full_temp["Stock"], errors='coerce').fillna(0)
        df_full_temp["Tarif d'achat"] = pd.to_numeric(df_full_temp["Tarif d'achat"], errors='coerce').fillna(0)
        df_full_temp["Conditionnement"] = pd.to_numeric(df_full_temp["Conditionnement"], errors='coerce').fillna(1)
        df_full_temp["Conditionnement"] = df_full_temp["Conditionnement"].apply(lambda x: int(x) if x > 0 else 1)
        # Ensure key textual columns are strings
        for str_col in ["Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article"]:
            df_full_temp[str_col] = df_full_temp[str_col].astype(str).str.strip()


        st.session_state.df_full = df_full_temp
        st.success("✅ Onglet 'Tableau final' lu et traité.")

        # --- Read 'Minimum de commande' sheet ---
        st.info("Lecture de l'onglet 'Minimum de commande'...")
        df_min_commande_temp = safe_read_excel(excel_file_buffer, sheet_name="Minimum de commande")
        
        min_order_dict_temp = {}
        if df_min_commande_temp is not None:
            supplier_col_name = "Fournisseur" # Expected column name for supplier
            min_amount_col_name = "Minimum de Commande" # Expected column name for min amount
            
            if supplier_col_name in df_min_commande_temp.columns and min_amount_col_name in df_min_commande_temp.columns:
                try:
                    # Clean supplier names and ensure min amount is numeric
                    df_min_commande_temp[supplier_col_name] = df_min_commande_temp[supplier_col_name].astype(str).str.strip()
                    df_min_commande_temp[min_amount_col_name] = pd.to_numeric(df_min_commande_temp[min_amount_col_name], errors='coerce')
                    
                    # Create dictionary, dropping rows where essential info is missing
                    min_order_dict_temp = df_min_commande_temp.dropna(
                        subset=[supplier_col_name, min_amount_col_name]
                    ).set_index(supplier_col_name)[min_amount_col_name].to_dict()
                    st.success(f"✅ Onglet 'Minimum de commande' lu. {len(min_order_dict_temp)} minimums chargés.")
                except Exception as e_min_proc:
                    st.error(f"❌ Erreur lors du traitement de l'onglet 'Minimum de commande': {e_min_proc}")
                    logging.error(f"Error processing 'Minimum de commande': {e_min_proc}")
            else:
                st.warning(f"⚠️ Colonnes '{supplier_col_name}' et/ou '{min_amount_col_name}' non trouvées dans 'Minimum de commande'. Les minimums de commande ne seront pas appliqués.")
        else:
            st.info("Onglet 'Minimum de commande' non trouvé ou vide. Aucun minimum de commande ne sera appliqué.")
        st.session_state.min_order_dict = min_order_dict_temp

        # --- Initial filtering and setup from df_full ---
        df_loaded = st.session_state.df_full
        try:
            # Filter out rows with missing critical identifiers early
            # Using .copy() to avoid SettingWithCopyWarning on df_init_filtered later
            df_init_filtered_temp = df_loaded[
                (df_loaded["Fournisseur"].notna()) & (df_loaded["Fournisseur"] != "") & (df_loaded["Fournisseur"] != "#FILTER") &
                (df_loaded["AF_RefFourniss"].notna()) & (df_loaded["AF_RefFourniss"] != "")
            ].copy()
            
            st.session_state.df_initial_filtered = df_init_filtered_temp

            # Identify sales week columns (heuristic based on position and numeric type)
            # This assumes sales data starts after a fixed number of initial info columns.
            first_potential_week_col_index = 12 # Example: if first 12 cols are info
            
            potential_sales_cols = []
            if len(df_loaded.columns) > first_potential_week_col_index:
                candidate_cols = df_loaded.columns[first_potential_week_col_index:].tolist()
                # Define columns that are definitely NOT sales weeks, even if they are numeric and appear late
                # This list might need adjustment based on actual file structures
                known_non_week_cols = [
                    "Tarif d'achat", "Conditionnement", "Stock", "Total", "Stock à terme", 
                    "Ventes N-1", "Ventes 12 semaines identiques N-1", 
                    "Ventes 12 dernières semaines", "Quantité à commander",
                    # Add any other known calculated/summary columns that might appear after info cols
                ] 
                # Also exclude known ID columns again, just in case they are past index 12
                known_id_cols = ["Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article"]
                
                exclude_from_sales_weeks = set(known_non_week_cols + known_id_cols)

                for col in candidate_cols:
                    if col not in exclude_from_sales_weeks:
                        # Check if the column in the *original* df_full is numeric-like
                        # Using df_loaded.get(col) for safety, though col comes from df_loaded.columns
                        if pd.api.types.is_numeric_dtype(df_loaded.get(col, pd.Series(dtype=object)).dtype):
                            potential_sales_cols.append(col)
                        # Optional: Add regex check for column name format e.g., YYYYWW or YYYYSWW
                        # elif re.match(r"(\d{4})S?(\d{1,2})", str(col), re.IGNORECASE):
                        #    potential_sales_cols.append(col) # Assume these are sales if they match format
            
            st.session_state.all_available_semaine_columns = potential_sales_cols
            if not potential_sales_cols:
                st.warning("⚠️ Aucune colonne de vente numérique n'a été automatiquement identifiée après les colonnes d'information initiales. Les calculs basés sur l'historique des ventes pourraient ne pas fonctionner.")
                logging.warning("No sales week columns identified based on current heuristics.")
            else:
                logging.info(f"Identified {len(potential_sales_cols)} potential sales week columns.")

            # Populate unique suppliers list from the filtered data
            if not df_init_filtered_temp.empty and "Fournisseur" in df_init_filtered_temp.columns:
                st.session_state.unique_suppliers_list = sorted(df_init_filtered_temp["Fournisseur"].unique().tolist())
            
            st.rerun() # Rerun to update UI with loaded data and clear the "upload file" state visually

        except KeyError as e_filter_key:
            st.error(f"❌ Erreur de clé (colonne manquante probable) lors du filtrage initial des données : '{e_filter_key}'.")
            logging.error(f"KeyError during initial filtering: {e_filter_key}")
            st.stop()
        except Exception as e_filter_other:
            st.error(f"❌ Erreur inattendue lors du filtrage initial des données : {e_filter_other}")
            logging.exception("Exception during initial filtering:")
            st.stop()
            
    except Exception as e_load_main:
        st.error(f"❌ Une erreur majeure est survenue lors du chargement ou du traitement initial du fichier : {e_load_main}")
        logging.exception("Major file loading/processing error:")
        # Clear df_full to allow re-upload attempt without full page reload
        st.session_state.df_full = None 
        st.stop()


# --- Main Application UI ---
# This part runs if data has been successfully loaded and initially processed
# (i.e., df_initial_filtered is available in session_state and is a DataFrame)
if 'df_initial_filtered' in st.session_state and isinstance(st.session_state.df_initial_filtered, pd.DataFrame):

    # Retrieve frequently used data from session state for easier access
    # df_full_data = st.session_state.df_full # The full, minimally processed data
    df_base_for_tabs = st.session_state.df_initial_filtered # Base data for tab-specific filtering
    
    all_suppliers_from_data = st.session_state.unique_suppliers_list
    min_order_amounts = st.session_state.min_order_dict
    identified_semaine_cols = st.session_state.all_available_semaine_columns

    # --- Define Tabs ---
    tab_titles = ["Prévision Commande", "Analyse Rotation Stock", "Vérification Stock", "Simulation Forecast"]
    tab1, tab2, tab3, tab4 = st.tabs(tab_titles)

    # ========================= TAB 1: Prévision Commande =========================
    with tab1:
        st.header("Prévision des Quantités à Commander")

        selected_fournisseurs_tab1 = render_supplier_checkboxes(
            "tab1", all_suppliers_from_data, default_select_all=True
        )

        df_display_tab1 = pd.DataFrame() # Initialize as empty
        if selected_fournisseurs_tab1:
            if not df_base_for_tabs.empty:
                df_display_tab1 = df_base_for_tabs[
                    df_base_for_tabs["Fournisseur"].isin(selected_fournisseurs_tab1)
                ].copy() # Use .copy() for any modifications
                st.caption(f"{len(df_display_tab1)} articles pour {len(selected_fournisseurs_tab1)} fournisseur(s) sélectionné(s).")
            else:
                 st.caption("Aucune donnée de base à filtrer pour les fournisseurs.")
        else:
            st.info("Veuillez sélectionner au moins un fournisseur.")

        st.markdown("---")

        if df_display_tab1.empty and selected_fournisseurs_tab1 : # Suppliers selected but no data for them
            st.warning("Aucun article trouvé pour le(s) fournisseur(s) sélectionné(s) dans les données de base.")
        elif not identified_semaine_cols and not df_display_tab1.empty:
            st.warning("Impossible de calculer : Aucune colonne de ventes (semaines) n'a été identifiée dans le fichier.")
        elif not df_display_tab1.empty : # Proceed if data is available for selected suppliers
            st.markdown("#### Paramètres de Calcul de Commande")
            col1_params_cmd, col2_params_cmd = st.columns(2)
            with col1_params_cmd:
                duree_couverture_semaines_cmd = st.number_input(
                    label="⏳ Durée de couverture souhaitée (en semaines)", 
                    min_value=1, max_value=260, value=4, step=1, key="duree_couv_cmd_ui"
                )
            with col2_params_cmd:
                montant_min_global_cmd = st.number_input(
                    label="💶 Montant minimum global de commande (€)", 
                    min_value=0.0, value=0.0, step=50.0, format="%.2f", key="montant_min_cmd_ui"
                )
            
            if st.button("🚀 Calculer les Quantités à Commander", key="calc_qte_cmd_btn"):
                with st.spinner("Calcul des quantités en cours..."):
                    # Pass only relevant data for calculation
                    result_tuple_cmd = calculer_quantite_a_commander(
                        df_display_tab1, 
                        identified_semaine_cols, 
                        montant_min_global_cmd, 
                        duree_couverture_semaines_cmd
                    )
                
                if result_tuple_cmd:
                    st.success("✅ Calcul des quantités terminé.")
                    quantites_calculees, ventes_n1_calc, ventes_12_n1_sim_calc, ventes_12_dern_calc, montant_total_calc = result_tuple_cmd
                    
                    df_result_cmd = df_display_tab1.copy() # Start with the data that was input to calculation
                    df_result_cmd["Qte Cmdée"] = quantites_calculees
                    df_result_cmd["Vts N-1 Total (calc)"] = ventes_n1_calc
                    df_result_cmd["Vts 12 N-1 Sim (calc)"] = ventes_12_n1_sim_calc # N-1 "same period"
                    df_result_cmd["Vts 12 Dern. (calc)"] = ventes_12_dern_calc # N "last 12"
                    
                    # Ensure 'Tarif d'achat' is numeric for total calculation; should be already from load
                    df_result_cmd["Tarif Ach."] = pd.to_numeric(df_result_cmd["Tarif d'achat"], errors='coerce').fillna(0)
                    df_result_cmd["Total Cmd (€)"] = df_result_cmd["Tarif Ach."] * df_result_cmd["Qte Cmdée"]
                    df_result_cmd["Stock Terme"] = df_result_cmd["Stock"] + df_result_cmd["Qte Cmdée"]
                    
                    # Store results in session state
                    st.session_state.commande_result_df = df_result_cmd
                    st.session_state.commande_calculated_total_amount = montant_total_calc
                    st.session_state.commande_suppliers_calculated_for = selected_fournisseurs_tab1 # Record which suppliers this calc was for
                    st.rerun() # Rerun to display results
                else:
                    st.error("❌ Le calcul des quantités a échoué ou n'a retourné aucun résultat.")

            # --- Display Commande Results ---
            if st.session_state.commande_result_df is not None:
                # Check if current UI selections match the selections for which results were calculated
                if st.session_state.commande_suppliers_calculated_for == selected_fournisseurs_tab1:
                    st.markdown("---")
                    st.markdown("#### Résultats de la Prévision de Commande")
                    
                    df_to_display_cmd = st.session_state.commande_result_df
                    calculated_total_cmd = st.session_state.commande_calculated_total_amount
                    suppliers_in_calc_cmd = st.session_state.commande_suppliers_calculated_for

                    st.metric(label="💰 Montant Total Commandé (calculé)", value=f"{calculated_total_cmd:,.2f} €")

                    # Warning for minimum order not met (if single supplier selected and has a minimum)
                    if len(suppliers_in_calc_cmd) == 1:
                        single_supplier_name = suppliers_in_calc_cmd[0]
                        if single_supplier_name in min_order_amounts:
                            required_min_for_supplier = min_order_amounts[single_supplier_name]
                            # Calculate actual total for this single supplier from the displayed results
                            actual_total_for_supplier = df_to_display_cmd[
                                df_to_display_cmd["Fournisseur"] == single_supplier_name
                            ]["Total Cmd (€)"].sum()
                                
                            if required_min_for_supplier > 0 and actual_total_for_supplier < required_min_for_supplier:
                                difference = required_min_for_supplier - actual_total_for_supplier
                                st.warning(
                                    f"⚠️ Minimum de commande non atteint pour {single_supplier_name}.\n"
                                    f"Montant actuel: **{actual_total_for_supplier:,.2f}€** | Requis: **{required_min_for_supplier:,.2f}€** "
                                    f"(Manque: {difference:,.2f}€)"
                                )
                    
                    # Define columns to display for the command results
                    cols_to_show_cmd = [
                        "Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article", 
                        "Stock", "Vts N-1 Total (calc)", "Vts 12 N-1 Sim (calc)", "Vts 12 Dern. (calc)",
                        "Conditionnement", "Qte Cmdée", "Stock Terme", "Tarif Ach.", "Total Cmd (€)"
                    ]
                    displayable_cols_cmd = [col for col in cols_to_show_cmd if col in df_to_display_cmd.columns]
                    
                    if not displayable_cols_cmd:
                        st.error("Aucune colonne de résultat à afficher pour la commande.")
                    else:
                        formatters_cmd = {
                            "Tarif Ach.": "{:,.2f}€", "Total Cmd (€)": "{:,.2f}€",
                            "Vts N-1 Total (calc)": "{:,.0f}", "Vts 12 N-1 Sim (calc)": "{:,.0f}",
                            "Vts 12 Dern. (calc)": "{:,.0f}", "Stock": "{:,.0f}",
                            "Conditionnement": "{:,.0f}", "Qte Cmdée": "{:,.0f}", "Stock Terme": "{:,.0f}"
                        }
                        st.dataframe(
                            df_to_display_cmd[displayable_cols_cmd].style.format(formatters_cmd, na_rep="-", thousands=",")
                        )

                    # --- Export Commande Logic ---
                    st.markdown("#### Exporter les Commandes Calculées")
                    # Filter for items with quantity ordered > 0
                    df_export_cmd_base = df_to_display_cmd[df_to_display_cmd["Qte Cmdée"] > 0].copy()
                    
                    if not df_export_cmd_base.empty:
                        excel_output_buffer_cmd = io.BytesIO()
                        sheets_created_cmd = 0
                        try:
                            with pd.ExcelWriter(excel_output_buffer_cmd, engine="openpyxl") as writer_cmd:
                                # Define columns for export, excluding 'Fournisseur' as it's per sheet
                                export_cols_per_sheet_cmd = [
                                    c for c in displayable_cols_cmd if c != 'Fournisseur'
                                ]
                                # Check if necessary columns for formula exist
                                qty_col_name = "Qte Cmdée"
                                price_col_name = "Tarif Ach."
                                total_col_name = "Total Cmd (€)"
                                formula_possible = False
                                if all(c in export_cols_per_sheet_cmd for c in [qty_col_name, price_col_name, total_col_name]):
                                    try:
                                        # Get 1-based index for column letters
                                        qty_col_letter = get_column_letter(export_cols_per_sheet_cmd.index(qty_col_name) + 1)
                                        price_col_letter = get_column_letter(export_cols_per_sheet_cmd.index(price_col_name) + 1)
                                        total_col_letter = get_column_letter(export_cols_per_sheet_cmd.index(total_col_name) + 1)
                                        formula_possible = True
                                    except ValueError: # index() not found
                                        formula_possible = False 
                                        logging.warning("Could not find Qty/Price/Total columns for formula in export.")

                                for supplier_for_sheet in suppliers_in_calc_cmd: # Iterate over suppliers for whom calc was run
                                    df_supplier_sheet_data = df_export_cmd_base[
                                        df_export_cmd_base["Fournisseur"] == supplier_for_sheet
                                    ]
                                    
                                    if not df_supplier_sheet_data.empty:
                                        df_to_write_to_sheet = df_supplier_sheet_data[export_cols_per_sheet_cmd].copy()
                                        num_data_rows = len(df_to_write_to_sheet)
                                        
                                        # Add TOTAL row and MINIMUM REQUIRED row
                                        # Determine a label column (e.g., 'Désignation Article' or the first available)
                                        label_col_for_summary = "Désignation Article"
                                        if label_col_for_summary not in export_cols_per_sheet_cmd:
                                            label_col_for_summary = export_cols_per_sheet_cmd[1] if len(export_cols_per_sheet_cmd) > 1 else export_cols_per_sheet_cmd[0]

                                        total_value_for_sheet = df_to_write_to_sheet[total_col_name].sum()
                                        min_req_for_sheet = min_order_amounts.get(supplier_for_sheet, 0)
                                        min_req_display = f"{min_req_for_sheet:,.2f}€" if min_req_for_sheet > 0 else "N/A"

                                        summary_rows_data = [
                                            {label_col_for_summary: "TOTAL", total_col_name: total_value_for_sheet},
                                            {label_col_for_summary: "Minimum Requis Fournisseur", total_col_name: min_req_display}
                                        ]
                                        df_summary_rows = pd.DataFrame(summary_rows_data, columns=export_cols_per_sheet_cmd).fillna('')
                                        
                                        df_final_sheet = pd.concat([df_to_write_to_sheet, df_summary_rows], ignore_index=True)
                                        
                                        safe_sheet_name = sanitize_sheet_name(supplier_for_sheet)
                                        try:
                                            df_final_sheet.to_excel(writer_cmd, sheet_name=safe_sheet_name, index=False)
                                            worksheet = writer_cmd.sheets[safe_sheet_name]
                                            
                                            # Apply formulas if possible
                                            if formula_possible and num_data_rows > 0:
                                                # Formulas for each item's total (rows 2 to num_data_rows + 1, as header is row 1)
                                                for r_idx in range(2, num_data_rows + 2):
                                                    formula_str = f"={qty_col_letter}{r_idx}*{price_col_letter}{r_idx}"
                                                    cell_to_write = worksheet[f"{total_col_letter}{r_idx}"]
                                                    cell_to_write.value = formula_str
                                                    cell_to_write.number_format = '#,##0.00€' # Apply currency format
                                                
                                                # Formula for the grand TOTAL row (num_data_rows + 2 is the first summary row)
                                                total_formula_cell_loc = f"{total_col_letter}{num_data_rows + 2}"
                                                sum_formula_str = f"=SUM({total_col_letter}2:{total_col_letter}{num_data_rows + 1})"
                                                cell_grand_total = worksheet[total_formula_cell_loc]
                                                cell_grand_total.value = sum_formula_str
                                                cell_grand_total.number_format = '#,##0.00€'
                                            sheets_created_cmd += 1
                                        except Exception as e_sheet_write:
                                            logging.error(f"Erreur écriture feuille Excel '{safe_sheet_name}': {e_sheet_write}")
                                            st.error(f"Erreur lors de la création de la feuille pour {supplier_for_sheet}.")
                            if sheets_created_cmd > 0:
                                writer_cmd.save() # Close writer to finalize file
                                excel_output_buffer_cmd.seek(0)
                                export_filename_cmd = (
                                    f"commandes_{'multiples_fournisseurs' if len(suppliers_in_calc_cmd) > 1 else sanitize_sheet_name(suppliers_in_calc_cmd[0])}"
                                    f"_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx"
                                )
                                st.download_button(
                                    label=f"📥 Télécharger Commandes ({sheets_created_cmd} feuille(s))",
                                    data=excel_output_buffer_cmd,
                                    file_name=export_filename_cmd,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key="dl_cmd_export_btn"
                                )
                            else:
                                st.info("Aucune donnée de commande à exporter (quantités commandées pourraient être nulles).")
                                
                        except Exception as e_excel_writer:
                            logging.exception(f"Erreur ExcelWriter pour commandes: {e_excel_writer}")
                            st.error("Une erreur est survenue lors de la préparation du fichier Excel pour les commandes.")
                    else:
                        st.info("Aucun article avec une quantité commandée > 0 à exporter.")
                else: # Results in session state do not match current supplier selection
                    st.info("Les résultats de commande affichés précédemment sont invalidés car la sélection de fournisseurs a changé. Veuillez relancer le calcul.")
        # End of Tab 1 specific logic when df_display_tab1 is not empty

    # ====================== TAB 2: Analyse Rotation Stock ======================
    with tab2:
        st.header("Analyse de la Rotation des Stocks")

        selected_fournisseurs_tab2 = render_supplier_checkboxes(
            "tab2", all_suppliers_from_data, default_select_all=True
        )
        
        df_display_tab2 = pd.DataFrame()
        if selected_fournisseurs_tab2:
            if not df_base_for_tabs.empty:
                df_display_tab2 = df_base_for_tabs[
                    df_base_for_tabs["Fournisseur"].isin(selected_fournisseurs_tab2)
                ].copy()
                st.caption(f"{len(df_display_tab2)} articles pour {len(selected_fournisseurs_tab2)} fournisseur(s) sélectionné(s).")
            else:
                st.caption("Aucune donnée de base à filtrer.")
        else:
            st.info("Veuillez sélectionner au moins un fournisseur.")
        
        st.markdown("---")

        if df_display_tab2.empty and selected_fournisseurs_tab2:
            st.warning("Aucun article trouvé pour le(s) fournisseur(s) sélectionné(s).")
        elif not identified_semaine_cols and not df_display_tab2.empty:
            st.warning("Impossible d'analyser : Aucune colonne de ventes (semaines) n'a été identifiée.")
        elif not df_display_tab2.empty:
            st.markdown("#### Paramètres d'Analyse de Rotation")
            col1_rot_params, col2_rot_params = st.columns(2)
            with col1_rot_params:
                period_options_rot = {"12 dernières semaines": 12, "52 dernières semaines": 52, "Toutes les données disponibles": 0}
                selected_period_label_rot = st.selectbox(
                    "⏳ Période d'analyse des ventes:", options=period_options_rot.keys(), 
                    key="rot_analysis_period_selectbox"
                )
                selected_period_weeks_rot = period_options_rot[selected_period_label_rot]
            
            with col2_rot_params:
                st.markdown("##### Options d'Affichage des Résultats")
                # Retrieve from session state to persist user choice
                show_all_rot_ui = st.checkbox(
                    "Afficher tous les articles", 
                    value=st.session_state.show_all_rotation_data, 
                    key="show_all_rot_ui_cb"
                )
                st.session_state.show_all_rotation_data = show_all_rot_ui # Persist choice
                
                # Threshold input, disabled if "show all" is checked
                rotation_filter_threshold_ui = st.number_input(
                    "... ou afficher articles avec ventes mensuelles moyennes <", 
                    min_value=0.0, 
                    value=st.session_state.rotation_threshold_value, 
                    step=0.1, format="%.1f", 
                    key="rot_filter_threshold_ui_numin",
                    disabled=show_all_rot_ui
                )
                if not show_all_rot_ui: # Only update if not showing all, to persist the value
                    st.session_state.rotation_threshold_value = rotation_filter_threshold_ui

            if st.button("🔄 Analyser la Rotation des Stocks", key="analyze_rotation_btn"):
                with st.spinner("Analyse de la rotation en cours..."):
                    df_rotation_results = calculer_rotation_stock(
                        df_display_tab2, identified_semaine_cols, selected_period_weeks_rot
                    )
                
                if df_rotation_results is not None: # Check if calculation was successful (not None)
                    st.success("✅ Analyse de rotation terminée.")
                    st.session_state.rotation_result_df = df_rotation_results
                    st.session_state.rotation_analysis_period_label = selected_period_label_rot # Store for display
                    st.session_state.rotation_suppliers_calculated_for = selected_fournisseurs_tab2
                    st.rerun()
                else:
                    st.error("❌ L'analyse de rotation a échoué ou n'a pas produit de résultats.")

            # --- Display Rotation Results ---
            if st.session_state.rotation_result_df is not None:
                if st.session_state.rotation_suppliers_calculated_for == selected_fournisseurs_tab2:
                    st.markdown("---")
                    st.markdown(f"#### Résultats de l'Analyse de Rotation ({st.session_state.rotation_analysis_period_label})")
                    
                    df_rotation_output_base = st.session_state.rotation_result_df
                    current_filter_threshold = st.session_state.rotation_threshold_value
                    show_all_filter_active = st.session_state.show_all_rotation_data
                    
                    df_filtered_for_display_rot = pd.DataFrame()
                    monthly_sales_col_name_rot = "Ventes Moy Mensuel (Période)"

                    if df_rotation_output_base.empty:
                        st.info("Aucune donnée de rotation à afficher (résultat du calcul vide).")
                    elif show_all_filter_active:
                        df_filtered_for_display_rot = df_rotation_output_base.copy()
                        st.caption(f"Affichage de {len(df_filtered_for_display_rot)} articles (tous les articles analysés).")
                    elif monthly_sales_col_name_rot in df_rotation_output_base.columns:
                        try:
                            # Ensure the filter column is numeric
                            sales_for_filter = pd.to_numeric(df_rotation_output_base[monthly_sales_col_name_rot], errors='coerce').fillna(0)
                            df_filtered_for_display_rot = df_rotation_output_base[sales_for_filter < current_filter_threshold].copy()
                            st.caption(
                                f"Filtré : Articles avec ventes moyennes < {current_filter_threshold:.1f}/mois. "
                                f"{len(df_filtered_for_display_rot)} / {len(df_rotation_output_base)} articles affichés."
                            )
                            if df_filtered_for_display_rot.empty:
                                st.info(f"Aucun article ne correspond au critère de filtre (ventes < {current_filter_threshold:.1f}/mois).")
                        except Exception as e_filter_rot:
                            st.error(f"Erreur lors du filtrage des résultats de rotation: {e_filter_rot}")
                            df_filtered_for_display_rot = df_rotation_output_base.copy() # Show all on error
                    else: # Filter column missing, show all
                        st.warning(f"Colonne '{monthly_sales_col_name_rot}' non trouvée pour le filtrage. Affichage de tous les articles.")
                        df_filtered_for_display_rot = df_rotation_output_base.copy()

                    if not df_filtered_for_display_rot.empty:
                        cols_to_show_rot = [
                            "Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article", 
                            "Tarif d'achat", "Stock", "Unités Vendues (Période)", 
                            "Ventes Moy Hebdo (Période)", "Ventes Moy Mensuel (Période)", 
                            "Semaines Stock (WoS)", "Rotation Unités (Proxy)", 
                            "Valeur Stock Actuel (€)", "COGS (Période)", "Rotation Valeur (Proxy)"
                        ]
                        # Filter for columns that actually exist in the dataframe
                        displayable_cols_rot = [col for col in cols_to_show_rot if col in df_filtered_for_display_rot.columns]
                        
                        # Round numeric columns for better display, handle inf
                        df_display_copy_rot = df_filtered_for_display_rot[displayable_cols_rot].copy()
                        numeric_cols_rounding_map_rot = {
                            "Tarif d'achat": 2, "Ventes Moy Hebdo (Période)": 2, 
                            "Ventes Moy Mensuel (Période)": 2, "Semaines Stock (WoS)": 1, 
                            "Rotation Unités (Proxy)": 2, "Valeur Stock Actuel (€)": 2, 
                            "COGS (Période)": 2, "Rotation Valeur (Proxy)": 2
                        }
                        for num_col, round_digits in numeric_cols_rounding_map_rot.items():
                            if num_col in df_display_copy_rot.columns:
                                df_display_copy_rot[num_col] = pd.to_numeric(df_display_copy_rot[num_col], errors='coerce')
                                if pd.api.types.is_numeric_dtype(df_display_copy_rot[num_col].dtype):
                                     df_display_copy_rot[num_col] = df_display_copy_rot[num_col].round(round_digits)
                        
                        df_display_copy_rot.replace([np.inf, -np.inf], 'Infini', inplace=True) # Replace inf with text for display

                        formatters_rot = {
                            "Tarif d'achat": "{:,.2f}€", "Stock": "{:,.0f}", 
                            "Unités Vendues (Période)": "{:,.0f}", 
                            "Ventes Moy Hebdo (Période)": "{:,.2f}", 
                            "Ventes Moy Mensuel (Période)": "{:,.2f}", 
                            "Semaines Stock (WoS)": "{}", # Already rounded or Infini
                            "Rotation Unités (Proxy)": "{}",
                            "Valeur Stock Actuel (€)": "{:,.2f}€", 
                            "COGS (Période)": "{:,.2f}€", 
                            "Rotation Valeur (Proxy)": "{}"
                        }
                        st.dataframe(df_display_copy_rot.style.format(formatters_rot, na_rep="-", thousands=","))

                        # --- Export Rotation Analysis ---
                        st.markdown("#### Exporter l'Analyse de Rotation Affichée")
                        excel_output_buffer_rot = io.BytesIO()
                        # Use df_display_copy_rot as it's already prepared (rounded, inf replaced)
                        df_export_rot = df_display_copy_rot # Already has the right columns and formatting
                        
                        sheet_name_label_rot = f"Rotation_{'Filtree' if not show_all_filter_active else 'Complete'}"
                        export_filename_base_rot = f"analyse_rotation_{'filtree' if not show_all_filter_active else 'complete'}"
                        
                        with pd.ExcelWriter(excel_output_buffer_rot, engine="openpyxl") as writer_rot:
                            df_export_rot.to_excel(writer_rot, sheet_name=sanitize_sheet_name(sheet_name_label_rot), index=False)
                        
                        excel_output_buffer_rot.seek(0)
                        current_suppliers_for_export_name = "multiples_fournisseurs"
                        if len(selected_fournisseurs_tab2) == 1:
                            current_suppliers_for_export_name = sanitize_sheet_name(selected_fournisseurs_tab2[0])
                        elif not selected_fournisseurs_tab2:
                             current_suppliers_for_export_name = "aucun_fournisseur"


                        export_filename_rot = (
                            f"{export_filename_base_rot}_{current_suppliers_for_export_name}"
                            f"_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx"
                        )
                        download_label_rot = f"📥 Télécharger Analyse {'Filtrée' if not show_all_filter_active else 'Complète'}"
                        st.download_button(
                            label=download_label_rot,
                            data=excel_output_buffer_rot,
                            file_name=export_filename_rot,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_rotation_export_btn"
                        )
                else: # Results in session state do not match current selection
                    st.info("Les résultats d'analyse de rotation affichés sont invalidés car la sélection de fournisseurs a changé. Veuillez relancer l'analyse.")
        # End of Tab 2 specific logic

    # ========================= TAB 3: Vérification Stock =========================
    with tab3:
        st.header("Vérification des Stocks Négatifs")
        st.caption("Cette analyse porte sur l'ensemble des articles du fichier 'Tableau final' (non filtré par fournisseur).")
        
        df_full_for_neg_check = st.session_state.get('df_full', None) # Use the complete df_full

        if df_full_for_neg_check is None or not isinstance(df_full_for_neg_check, pd.DataFrame):
            st.warning("Les données complètes ('Tableau final') n'ont pas été chargées. Impossible de vérifier les stocks négatifs.")
        elif df_full_for_neg_check.empty:
            st.info("Le 'Tableau final' est vide. Aucune vérification de stock négatif à effectuer.")
        else:
            stock_col_name = "Stock"
            if stock_col_name not in df_full_for_neg_check.columns:
                st.error(f"La colonne '{stock_col_name}' est introuvable dans 'Tableau final'. Vérification impossible.")
            else:
                # Stock column should already be numeric from initial load
                df_negative_stocks = df_full_for_neg_check[df_full_for_neg_check[stock_col_name] < 0].copy()
                
                if df_negative_stocks.empty:
                    st.success("✅ Aucun article avec un stock négatif n'a été trouvé dans le fichier.")
                else:
                    st.warning(f"⚠️ **{len(df_negative_stocks)} article(s) trouvé(s) avec un stock négatif !**")
                    
                    cols_to_show_neg_stock = ["Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article", "Stock"]
                    displayable_cols_neg_stock = [col for col in cols_to_show_neg_stock if col in df_negative_stocks.columns]
                    
                    if not displayable_cols_neg_stock:
                        st.error("Impossible d'afficher les détails des stocks négatifs (colonnes d'identification manquantes).")
                    else:
                        st.dataframe(
                            df_negative_stocks[displayable_cols_neg_stock].style.format(
                                {"Stock": "{:,.0f}"}, na_rep="-"
                            ).apply(
                                lambda s: ['background-color:#FADBD8' if s.name == stock_col_name and val < 0 else '' for val in s], 
                                axis=0 # Apply column-wise to target 'Stock' column by name
                            )
                        )
                        st.markdown("---")
                        st.markdown("#### Exporter la Liste des Stocks Négatifs")
                        excel_output_buffer_neg = io.BytesIO()
                        df_export_neg_stock = df_negative_stocks[displayable_cols_neg_stock].copy()
                        try:
                            with pd.ExcelWriter(excel_output_buffer_neg, engine="openpyxl") as writer_neg:
                                df_export_neg_stock.to_excel(writer_neg, sheet_name="Stocks_Negatifs", index=False)
                            
                            excel_output_buffer_neg.seek(0)
                            export_filename_neg = f"stocks_negatifs_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx"
                            st.download_button(
                                "📥 Télécharger la Liste des Stocks Négatifs",
                                data=excel_output_buffer_neg,
                                file_name=export_filename_neg,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="dl_neg_stock_export_btn"
                            )
                        except Exception as e_export_neg:
                            st.error(f"Erreur lors de la préparation de l'export des stocks négatifs: {e_export_neg}")
                            logging.exception("Error exporting negative stocks:")
    # End of Tab 3

    # ========================= TAB 4: Simulation Forecast =========================
    with tab4:
        st.header("Simulation de Forecast Annuel")

        selected_fournisseurs_tab4 = render_supplier_checkboxes(
            "tab4", all_suppliers_from_data, default_select_all=True
        )

        df_display_tab4 = pd.DataFrame()
        if selected_fournisseurs_tab4:
            if not df_base_for_tabs.empty:
                df_display_tab4 = df_base_for_tabs[
                    df_base_for_tabs["Fournisseur"].isin(selected_fournisseurs_tab4)
                ].copy()
                st.caption(f"{len(df_display_tab4)} articles pour {len(selected_fournisseurs_tab4)} fournisseur(s) sélectionné(s).")
            else:
                st.caption("Aucune donnée de base à filtrer.")
        else:
            st.info("Veuillez sélectionner au moins un fournisseur.")
        
        st.markdown("---")
        
        st.warning(
            "🚨 **Hypothèse importante :** La saisonnalité mensuelle est une approximation basée sur un découpage "
            "des 52 semaines de l'année N-1. La précision dépend de la régularité des ventes N-1."
        )

        if df_display_tab4.empty and selected_fournisseurs_tab4:
            st.warning("Aucun article trouvé pour le(s) fournisseur(s) sélectionné(s).")
        elif len(identified_semaine_cols) < 52 and not df_display_tab4.empty :
            st.warning(f"Données historiques insuffisantes ({len(identified_semaine_cols)} semaines trouvées). Une simulation basée sur N-1 complet (52 semaines) n'est pas possible.")
        elif not df_display_tab4.empty:
            st.markdown("#### Paramètres de Simulation de Forecast")
            
            all_calendar_months = list(calendar.month_name)[1:] # Jan to Dec
            # Persist selected months
            selected_months_for_forecast_ui = st.multiselect(
                "📅 Mois à inclure dans la simulation:", 
                options=all_calendar_months, 
                default=st.session_state.forecast_selected_months_ui, # Use persisted value
                key="fcst_months_multiselect_ui"
            )
            st.session_state.forecast_selected_months_ui = selected_months_for_forecast_ui # Persist

            sim_type_options_fcst = ('Simple Progression', 'Objectif Montant')
            selected_sim_type_fcst_ui = st.radio(
                "⚙️ Type de Simulation:", 
                options=sim_type_options_fcst, 
                horizontal=True,
                index=st.session_state.forecast_sim_type_radio_index, # Persist choice
                key="fcst_sim_type_radio_ui"
            )
            st.session_state.forecast_sim_type_radio_index = sim_type_options_fcst.index(selected_sim_type_fcst_ui) # Persist

            prog_percentage_fcst_ui = 0.0
            target_amount_fcst_ui = 0.0
            
            col1_fcst_params, col2_fcst_params = st.columns(2)
            with col1_fcst_params:
                if selected_sim_type_fcst_ui == 'Simple Progression':
                    prog_percentage_fcst_ui = st.number_input(
                        label="📈 Taux de Progression Annuel Cible (%) vs N-1", 
                        min_value=-100.0, 
                        value=st.session_state.forecast_progression_percentage_ui, # Persist
                        step=0.5, format="%.1f", 
                        key="fcst_prog_pct_ui_numin"
                    )
                    st.session_state.forecast_progression_percentage_ui = prog_percentage_fcst_ui # Persist
            with col2_fcst_params:
                if selected_sim_type_fcst_ui == 'Objectif Montant':
                    target_amount_fcst_ui = st.number_input(
                        label="🎯 Montant Objectif (€) pour la période sélectionnée", 
                        min_value=0.0, 
                        value=st.session_state.forecast_target_amount_ui, # Persist
                        step=1000.0, format="%.2f", 
                        key="fcst_target_amt_ui_numin"
                    )
                    st.session_state.forecast_target_amount_ui = target_amount_fcst_ui # Persist

            if st.button("▶️ Lancer la Simulation de Forecast", key="run_forecast_sim_btn"):
                if not selected_months_for_forecast_ui:
                    st.error("Veuillez sélectionner au moins un mois pour la simulation.")
                else:
                    with st.spinner("Simulation du forecast en cours..."):
                        df_forecast_sim_result, grand_total_sim_amount = calculer_forecast_simulation_v3(
                            df_display_tab4, 
                            identified_semaine_cols, 
                            selected_months_for_forecast_ui, 
                            selected_sim_type_fcst_ui, 
                            prog_percentage_fcst_ui, 
                            target_amount_fcst_ui
                        )
                    
                    if df_forecast_sim_result is not None:
                        st.success("✅ Simulation de forecast terminée.")
                        st.session_state.forecast_result_df = df_forecast_sim_result
                        st.session_state.forecast_grand_total_amount = grand_total_sim_amount
                        # Store parameters used for this calculation to check against UI changes
                        st.session_state.forecast_simulation_params_calculated_for = {
                            'suppliers': selected_fournisseurs_tab4,
                            'months': selected_months_for_forecast_ui,
                            'type': selected_sim_type_fcst_ui,
                            'prog_pct': prog_percentage_fcst_ui if selected_sim_type_fcst_ui == 'Simple Progression' else 0,
                            'obj_amt': target_amount_fcst_ui if selected_sim_type_fcst_ui == 'Objectif Montant' else 0,
                        }
                        st.rerun()
                    else:
                        st.error("❌ La simulation de forecast a échoué ou n'a pas produit de résultats.")
            
            # --- Display Forecast Simulation Results ---
            if st.session_state.forecast_result_df is not None:
                # Construct current UI parameters for comparison
                current_ui_params_fcst = {
                    'suppliers': selected_fournisseurs_tab4,
                    'months': selected_months_for_forecast_ui, # From UI widget
                    'type': selected_sim_type_fcst_ui, # From UI widget
                    'prog_pct': st.session_state.forecast_progression_percentage_ui if selected_sim_type_fcst_ui == 'Simple Progression' else 0,
                    'obj_amt': st.session_state.forecast_target_amount_ui if selected_sim_type_fcst_ui == 'Objectif Montant' else 0,
                }
                
                if st.session_state.forecast_simulation_params_calculated_for == current_ui_params_fcst:
                    st.markdown("---")
                    st.markdown("#### Résultats de la Simulation de Forecast")
                    
                    df_to_display_fcst = st.session_state.forecast_result_df
                    grand_total_fcst_disp = st.session_state.forecast_grand_total_amount
                    
                    if df_to_display_fcst.empty:
                        st.info("Aucun résultat de simulation à afficher (le DataFrame est vide).")
                    else:
                        # Columns should already be ordered by the calculation function
                        # Prepare formatters (these are illustrative, adapt as needed)
                        formatters_fcst = {"Tarif d'achat": "{:,.2f}€", "Conditionnement": "{:,.0f}"}
                        # Add formatters for dynamic N-1, Qty, Amount columns
                        for month_disp in selected_months_for_forecast_ui: # Use months from current UI selection
                            if f"Ventes N-1 {month_disp}" in df_to_display_fcst.columns:
                                formatters_fcst[f"Ventes N-1 {month_disp}"] = "{:,.0f}"
                            if f"Qté Prév. {month_disp}" in df_to_display_fcst.columns:
                                formatters_fcst[f"Qté Prév. {month_disp}"] = "{:,.0f}"
                            if f"Montant Prév. {month_disp} (€)" in df_to_display_fcst.columns:
                                formatters_fcst[f"Montant Prév. {month_disp} (€)"] = "{:,.2f}€"
                        # Totals
                        if "Vts N-1 Tot (Mois Sel.)" in df_to_display_fcst.columns: formatters_fcst["Vts N-1 Tot (Mois Sel.)"] = "{:,.0f}"
                        if "Qté Tot Prév (Mois Sel.)" in df_to_display_fcst.columns: formatters_fcst["Qté Tot Prév (Mois Sel.)"] = "{:,.0f}"
                        if "Mnt Tot Prév (€) (Mois Sel.)" in df_to_display_fcst.columns: formatters_fcst["Mnt Tot Prév (€) (Mois Sel.)"] = "{:,.2f}€"

                        try:
                            st.dataframe(df_to_display_fcst.style.format(formatters_fcst, na_rep="-", thousands=","))
                        except Exception as e_fmt_fcst:
                            st.error(f"Erreur de formatage lors de l'affichage des résultats du forecast: {e_fmt_fcst}")
                            st.dataframe(df_to_display_fcst) # Display raw on error

                        st.metric(label="💰 Montant Total Général Prévisionnel (€) pour les mois sélectionnés", value=f"{grand_total_fcst_disp:,.2f} €")

                        # --- Export Forecast Simulation Results ---
                        st.markdown("#### Exporter la Simulation de Forecast")
                        excel_output_buffer_fcst = io.BytesIO()
                        # df_to_display_fcst already contains the correctly ordered columns
                        df_export_fcst = df_to_display_fcst.copy()
                        try:
                            sim_type_filename_part = selected_sim_type_fcst_ui.replace(' ', '_').lower()
                            with pd.ExcelWriter(excel_output_buffer_fcst, engine="openpyxl") as writer_fcst:
                                df_export_fcst.to_excel(writer_fcst, sheet_name=sanitize_sheet_name(f"Forecast_{sim_type_filename_part}"), index=False)
                            
                            excel_output_buffer_fcst.seek(0)
                            
                            current_suppliers_for_export_name_fcst = "multiples_fournisseurs"
                            if len(selected_fournisseurs_tab4) == 1:
                                current_suppliers_for_export_name_fcst = sanitize_sheet_name(selected_fournisseurs_tab4[0])
                            elif not selected_fournisseurs_tab4:
                                 current_suppliers_for_export_name_fcst = "aucun_fournisseur"

                            export_filename_fcst = (
                                f"simulation_forecast_{sim_type_filename_part}_{current_suppliers_for_export_name_fcst}"
                                f"_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx"
                            )
                            st.download_button(
                                "📥 Télécharger la Simulation",
                                data=excel_output_buffer_fcst,
                                file_name=export_filename_fcst,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="dl_forecast_sim_export_btn"
                            )
                        except Exception as e_export_fcst:
                            st.error(f"Erreur lors de la préparation de l'export de la simulation: {e_export_fcst}")
                            logging.exception("Error exporting forecast simulation:")
                else: # Parameters changed since last calculation
                    st.info("Les résultats de simulation affichés sont invalidés car les paramètres ou la sélection de fournisseurs ont changé. Veuillez relancer la simulation.")
        # End of Tab 4 specific logic

# --- App Footer / Initial Message if no file is loaded ---
elif not uploaded_file: # This implies st.session_state.df_full is None from initialization or reset
    st.info("👋 Bienvenue ! Chargez votre fichier Excel principal pour démarrer l'analyse et les prévisions.")
    
    # Offer a way to reset the entire application state if needed
    if st.button("🔄 Réinitialiser l'Application (efface toutes les données en session)"):
        keys_to_clear_full_reset = list(st.session_state.keys())
        for key_to_del in keys_to_clear_full_reset:
            del st.session_state[key_to_del]
        # Re-initialize defaults after full clear might be good, or let Streamlit handle on rerun
        # For a full reset, clearing and rerunning is usually enough.
        logging.info("Application state fully reset by user.")
        st.rerun()

# Fallback if df_initial_filtered is somehow not a DataFrame but key exists (defensive)
elif 'df_initial_filtered' in st.session_state and not isinstance(st.session_state.df_initial_filtered, pd.DataFrame):
    st.error("Erreur interne : L'état des données filtrées est invalide. Veuillez recharger le fichier.")
    st.session_state.df_full = None # Force re-upload
    if st.button("Réessayer de charger"):
        st.rerun()
