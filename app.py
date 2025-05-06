import streamlit as st
import pandas as pd
import numpy as np
import io
import logging
import re
import openpyxl # Utilisé indirectement par pd.ExcelWriter(engine='openpyxl')
from openpyxl.utils import get_column_letter
import calendar
# import zipfile # Si vous décidez d'implémenter l'export ZIP plus tard

# --- Logging Configuration ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- Helper Functions ---

def safe_read_excel(uploaded_file, sheet_name, **kwargs):
    """ Safely reads an Excel sheet, returning None if sheet not found or error occurs. """
    try:
        if isinstance(uploaded_file, io.BytesIO): uploaded_file.seek(0)
        file_name = getattr(uploaded_file, 'name', '')
        engine = 'openpyxl' if file_name.lower().endswith('.xlsx') else None
        
        logging.debug(f"Attempting to read sheet: '{sheet_name}' with kwargs: {kwargs}")
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, engine=engine, **kwargs)
        logging.debug(f"Read sheet '{sheet_name}'. DataFrame empty: {df.empty}, Columns: {df.columns.tolist()}, Shape: {df.shape}")
        
        # Si header est spécifié, pandas peut lire des colonnes même si toutes les lignes de données sont vides.
        # Un DataFrame avec des colonnes mais aucune ligne est df.empty == True.
        # On considère l'onglet "problématique" si aucune colonne n'est lue (souvent si onglet non trouvé ou totalement vide)
        if len(df.columns) == 0: #  and df.empty (implicite si pas de colonnes)
             logging.warning(f"Sheet '{sheet_name}' was read but has no columns (likely empty or not found as expected).")
             # Le message à l'utilisateur est géré par l'appelant généralement
             return None
        return df
    except ValueError as e:
        if f"Worksheet named '{sheet_name}' not found" in str(e) or f"'{sheet_name}' not found" in str(e):
             logging.warning(f"Sheet '{sheet_name}' not found in the Excel file.")
             st.warning(f"⚠️ Onglet '{sheet_name}' non trouvé dans le fichier Excel.")
        else:
             logging.error(f"ValueError reading sheet '{sheet_name}': {e}")
             st.error(f"❌ Erreur de valeur lors de la lecture de l'onglet '{sheet_name}': {e}.")
        return None
    except FileNotFoundError: # Should not happen with BytesIO
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

        qpond = (0.5 * avg12last + 0.2 * avg12N1 + 0.3 * avg12N1s)
        qnec = qpond * duree_semaines
        
        qcomm_series = (qnec - df_calc["Stock"]).apply(lambda x: max(0, x))
        
        cond = df_calc["Conditionnement"]
        stock = df_calc["Stock"]
        tarif = df_calc["Tarif d'achat"]
        
        qcomm = qcomm_series.tolist()

        for i in range(len(qcomm)):
            c = cond.iloc[i]
            q = qcomm[i]
            if q > 0 and c > 0:
                qcomm[i] = int(np.ceil(q / c) * c)
            elif q > 0 and c <= 0:
                # Log this occurrence for audit, user facing message might be too verbose if many such items
                logging.warning(f"Article index {df_calc.index[i]} (Ref: {df_calc.get('Référence Article', pd.Series(['N/A']))[i]}) Qté {q:.2f} ignorée car conditionnement est {c}.")
                qcomm[i] = 0 
            else: # q <= 0
                qcomm[i] = 0
        
        if nb_semaines_recentes > 0:
            for i in range(len(qcomm)):
                c = cond.iloc[i]
                vr_count = (df_calc[semaine_columns[-nb_semaines_recentes:]].iloc[i] > 0).sum()
                if vr_count >= 2 and stock.iloc[i] <= 1 and c > 0:
                    qcomm[i] = max(qcomm[i], c)

        for i in range(len(qcomm)):
            vt_n1_item = ventes_N1.iloc[i]
            vr_sum_item = v12last.iloc[i]
            if vt_n1_item < 6 and vr_sum_item < 2:
                qcomm[i] = 0

        qcomm_df_temp = pd.Series(qcomm, index=df_calc.index)
        mt_avant_ajustement = (qcomm_df_temp * tarif).sum()

        if montant_minimum_input > 0 and mt_avant_ajustement < montant_minimum_input:
            mt_actuel = mt_avant_ajustement
            eligible_for_increment = []
            for i in range(len(qcomm)):
                if qcomm[i] > 0 and cond.iloc[i] > 0 and tarif.iloc[i] > 0:
                    eligible_for_increment.append(i)

            if not eligible_for_increment:
                if mt_actuel < montant_minimum_input:
                    st.warning(
                        f"Impossible d'atteindre le montant minimum de {montant_minimum_input:,.2f}€. "
                        f"Montant actuel: {mt_actuel:,.2f}€. "
                        "Aucun article commandé avec conditionnement et tarif valides pour incrémentation."
                    )
            else:
                idx_ptr_eligible = 0
                max_iter_loop = len(eligible_for_increment) * 20 + 1 
                iters = 0
                while mt_actuel < montant_minimum_input and iters < max_iter_loop:
                    iters += 1
                    original_df_idx = eligible_for_increment[idx_ptr_eligible]
                    c_item = cond.iloc[original_df_idx]
                    p_item = tarif.iloc[original_df_idx]
                    
                    qcomm[original_df_idx] += c_item
                    mt_actuel += c_item * p_item
                    
                    idx_ptr_eligible = (idx_ptr_eligible + 1) % len(eligible_for_increment)
                
                if iters >= max_iter_loop and mt_actuel < montant_minimum_input:
                    st.error(
                        f"Ajustement du montant minimum : Nombre maximum d'itérations ({max_iter_loop}) atteint. "
                        f"Montant actuel: {mt_actuel:,.2f}€ / Requis: {montant_minimum_input:,.2f}€. "
                    )
        
        qcomm_final_series = pd.Series(qcomm, index=df_calc.index)
        mt_final = (qcomm_final_series * tarif).sum()
        
        return (qcomm, ventes_N1, v12N1, v12last, mt_final)

    except KeyError as e:
        st.error(f"Erreur de clé (colonne manquante probable) lors du calcul des quantités : '{e}'.")
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
            return pd.DataFrame()

        required_cols = ["Stock", "Tarif d'achat"]
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            st.error(f"Colonnes manquantes pour l'analyse de rotation : {', '.join(missing_cols)}")
            return None

        df_rotation = df.copy()

        if periode_semaines and periode_semaines > 0 and len(semaine_columns) >= periode_semaines:
            semaines_analyse = semaine_columns[-periode_semaines:]
            nb_semaines_analyse = periode_semaines
        elif periode_semaines and periode_semaines > 0: # Not enough history
            semaines_analyse = semaine_columns
            nb_semaines_analyse = len(semaine_columns)
            st.caption(f"Période d'analyse ajustée à {nb_semaines_analyse} semaines (données disponibles).")
        else: # periode_semaines is 0 or invalid, use all available
            semaines_analyse = semaine_columns
            nb_semaines_analyse = len(semaine_columns)
        
        if not semaines_analyse:
            st.warning("Aucune colonne de vente disponible pour l'analyse de rotation.")
            metric_cols = ["Unités Vendues (Période)", "Ventes Moy Hebdo (Période)", "Ventes Moy Mensuel (Période)",
                           "Semaines Stock (WoS)", "Rotation Unités (Proxy)", "COGS (Période)", 
                           "Valeur Stock Actuel (€)", "Rotation Valeur (Proxy)"]
            for m_col in metric_cols: df_rotation[m_col] = 0.0
            return df_rotation

        for col in semaines_analyse:
            df_rotation[col] = pd.to_numeric(df_rotation[col], errors='coerce').fillna(0)

        df_rotation["Unités Vendues (Période)"] = df_rotation[semaines_analyse].sum(axis=1)
        
        df_rotation["Ventes Moy Hebdo (Période)"] = df_rotation["Unités Vendues (Période)"] / nb_semaines_analyse if nb_semaines_analyse > 0 else 0.0
            
        avg_weeks_per_month = 52 / 12.0
        df_rotation["Ventes Moy Mensuel (Période)"] = df_rotation["Ventes Moy Hebdo (Période)"] * avg_weeks_per_month
        
        df_rotation["Stock"] = pd.to_numeric(df_rotation["Stock"], errors='coerce').fillna(0)
        df_rotation["Tarif d'achat"] = pd.to_numeric(df_rotation["Tarif d'achat"], errors='coerce').fillna(0)
        
        denom_wos = df_rotation["Ventes Moy Hebdo (Période)"]
        df_rotation["Semaines Stock (WoS)"] = np.divide(
            df_rotation["Stock"], denom_wos, 
            out=np.full_like(df_rotation["Stock"], np.inf, dtype=np.float64),
            where=denom_wos != 0
        )
        df_rotation.loc[df_rotation["Stock"] <= 0, "Semaines Stock (WoS)"] = 0.0

        denom_rot_unit = df_rotation["Stock"]
        df_rotation["Rotation Unités (Proxy)"] = np.divide(
            df_rotation["Unités Vendues (Période)"], denom_rot_unit,
            out=np.full_like(denom_rot_unit, np.inf, dtype=np.float64),
            where=denom_rot_unit != 0
        )
        df_rotation.loc[(df_rotation["Unités Vendues (Période)"] <= 0) & (denom_rot_unit <= 0), "Rotation Unités (Proxy)"] = 0.0
        df_rotation.loc[(df_rotation["Unités Vendues (Période)"] <= 0) & (denom_rot_unit > 0), "Rotation Unités (Proxy)"] = 0.0

        df_rotation["COGS (Période)"] = df_rotation["Unités Vendues (Période)"] * df_rotation["Tarif d'achat"]
        df_rotation["Valeur Stock Actuel (€)"] = df_rotation["Stock"] * df_rotation["Tarif d'achat"]
        
        denom_rot_val = df_rotation["Valeur Stock Actuel (€)"]
        df_rotation["Rotation Valeur (Proxy)"] = np.divide(
            df_rotation["COGS (Période)"], denom_rot_val,
            out=np.full_like(denom_rot_val, np.inf, dtype=np.float64),
            where=denom_rot_val != 0
        )
        df_rotation.loc[(df_rotation["COGS (Période)"] <= 0) & (denom_rot_val <= 0), "Rotation Valeur (Proxy)"] = 0.0
        df_rotation.loc[(df_rotation["COGS (Période)"] <= 0) & (denom_rot_val > 0), "Rotation Valeur (Proxy)"] = 0.0

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
        logging.warning(f"approx_weeks_to_months expects 52 columns, got {len(week_columns_52) if week_columns_52 else 0}. Returning empty map.")
        return month_map

    weeks_per_month_approx = 52 / 12.0
    
    for i in range(1, 13): # For months 1 to 12
        month_name = calendar.month_name[i]
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

        required_cols = ["Référence Article", "Désignation Article", "Conditionnement", "Tarif d'achat", "Fournisseur"]
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            st.error(f"Colonnes manquantes pour la simulation : {', '.join(missing_cols)}")
            return None, 0.0

        years_in_cols = set()
        parsed_week_cols = []
        for col_name in all_semaine_columns:
            if isinstance(col_name, str):
                match = re.match(r"(\d{4})S?(\d{1,2})", col_name, re.IGNORECASE)
                if match:
                    year = int(match.group(1))
                    week = int(match.group(2))
                    if 1 <= week <= 53:
                        years_in_cols.add(year)
                        parsed_week_cols.append({'year': year, 'week': week, 'col': col_name, 'sort_key': year * 100 + week})
        
        if not years_in_cols:
            st.error("Impossible de déterminer les années à partir des noms de colonnes semaines. Format attendu 'YYYYWW' ou 'YYYYSwW'.")
            return None, 0.0

        parsed_week_cols.sort(key=lambda x: x['sort_key'])
        
        year_n = max(years_in_cols) if years_in_cols else 0 # Current year estimation
        year_n_minus_1 = year_n - 1
        
        st.caption(f"Simulation basée sur N-1 (Année N détectée: {year_n}, Année N-1 utilisée: {year_n_minus_1})")

        n1_week_cols_data = [item for item in parsed_week_cols if item['year'] == year_n_minus_1]
        
        if len(n1_week_cols_data) < 52:
            st.error(f"Données N-1 ({year_n_minus_1}) insuffisantes. {len(n1_week_cols_data)} semaines trouvées, 52 requises.")
            return None, 0.0
        
        n1_week_cols_for_mapping = [item['col'] for item in n1_week_cols_data[:52]] # Use first 52 weeks of N-1

        df_sim = df[required_cols].copy()
        df_sim["Tarif d'achat"] = pd.to_numeric(df_sim["Tarif d'achat"], errors='coerce').fillna(0)
        df_sim["Conditionnement"] = pd.to_numeric(df_sim["Conditionnement"], errors='coerce').fillna(1)
        df_sim["Conditionnement"] = df_sim["Conditionnement"].apply(lambda x: 1 if x <= 0 else int(x))

        missing_n1_in_df = [col for col in n1_week_cols_for_mapping if col not in df.columns]
        if missing_n1_in_df:
            st.error(f"Erreur interne: Colonnes N-1 mappées ({', '.join(missing_n1_in_df)}) non trouvées dans les données de vente du DataFrame.")
            return None, 0.0
            
        df_n1_sales_data = df[n1_week_cols_for_mapping].copy()
        for col in n1_week_cols_for_mapping:
            df_n1_sales_data[col] = pd.to_numeric(df_n1_sales_data[col], errors='coerce').fillna(0)

        month_col_map_n1 = approx_weeks_to_months(n1_week_cols_for_mapping)
        
        total_n1_sales_selected_months_series = pd.Series(0.0, index=df_sim.index)
        monthly_sales_n1_for_selected_months = {}

        for month_name in selected_months:
            sales_this_month_n1 = pd.Series(0.0, index=df_sim.index)
            if month_name in month_col_map_n1 and month_col_map_n1[month_name]:
                actual_cols_for_month_n1 = [col for col in month_col_map_n1[month_name] if col in df_n1_sales_data.columns]
                if actual_cols_for_month_n1:
                    sales_this_month_n1 = df_n1_sales_data[actual_cols_for_month_n1].sum(axis=1)
            
            monthly_sales_n1_for_selected_months[month_name] = sales_this_month_n1
            total_n1_sales_selected_months_series += sales_this_month_n1
            df_sim[f"Ventes N-1 {month_name}"] = sales_this_month_n1 # Store N-1 sales for this month in output
        
        df_sim["Vts N-1 Tot (Mois Sel.)"] = total_n1_sales_selected_months_series

        period_seasonality_factors = {}
        safe_total_n1_for_selected_months = total_n1_sales_selected_months_series.copy()

        for month_name in selected_months:
            month_sales_n1 = monthly_sales_n1_for_selected_months.get(month_name, pd.Series(0.0, index=df_sim.index))
            # Calculate seasonality factor: (Month's N-1 Sales) / (Total N-1 Sales for ALL Selected Months)
            factor = np.divide(month_sales_n1, safe_total_n1_for_selected_months, 
                               out=np.zeros_like(month_sales_n1, dtype=float), # Output 0 where denom is 0
                               where=safe_total_n1_for_selected_months!=0)
            period_seasonality_factors[month_name] = pd.Series(factor, index=df_sim.index).fillna(0)

        base_monthly_forecast_qty_map = {}

        if sim_type == 'Simple Progression':
            prog_factor = 1 + (progression_pct / 100.0)
            total_forecast_qty_for_selected_period = total_n1_sales_selected_months_series * prog_factor
            
            for month_name in selected_months:
                seasonality_for_month = period_seasonality_factors.get(month_name, pd.Series(0.0, index=df_sim.index))
                base_monthly_forecast_qty_map[month_name] = total_forecast_qty_for_selected_period * seasonality_for_month
        
        elif sim_type == 'Objectif Montant':
            if objectif_montant <= 0:
                st.error("Pour 'Objectif Montant', l'objectif doit être supérieur à 0.")
                return None, 0.0

            total_n1_sales_units_all_items_for_period = total_n1_sales_selected_months_series.sum() # Sum of all N-1 units for all items in selected period

            if total_n1_sales_units_all_items_for_period <= 0: # N-1 sales are effectively zero for ALL items in the selected period
                st.warning(
                    "Les ventes N-1 pour les mois sélectionnés sont nulles globalement. "
                    "Tentative de répartition égale du montant objectif sur les mois, puis sur les articles (avec tarif > 0)."
                )
                num_selected_months = len(selected_months)
                if num_selected_months == 0: return None, 0.0
                
                target_amount_per_month_overall = objectif_montant / num_selected_months
                num_items_with_positive_price = (df_sim["Tarif d'achat"] > 0).sum()

                for month_name in selected_months:
                    if num_items_with_positive_price == 0: # No items to assign quantity to
                        base_monthly_forecast_qty_map[month_name] = pd.Series(0.0, index=df_sim.index)
                    else:
                        # Distribute this month's overall target amount equally among items that have a price
                        target_amount_per_item_this_month = target_amount_per_month_overall / num_items_with_positive_price
                        base_monthly_forecast_qty_map[month_name] = np.divide(
                            target_amount_per_item_this_month, df_sim["Tarif d'achat"],
                            out=np.zeros_like(df_sim["Tarif d'achat"], dtype=float),
                            where=df_sim["Tarif d'achat"] != 0 # Qty is 0 if price is 0
                        )
            else: # Normal case: N-1 sales exist for at least some items in the period
                for month_name in selected_months:
                    seasonality_for_month = period_seasonality_factors.get(month_name, pd.Series(0.0, index=df_sim.index))
                    # Target amount for this specific month for each item (vectorized by seasonality)
                    target_amount_for_this_month_per_item = objectif_montant * seasonality_for_month
                    
                    base_monthly_forecast_qty_map[month_name] = np.divide(
                        target_amount_for_this_month_per_item, df_sim["Tarif d'achat"],
                        out=np.zeros_like(df_sim["Tarif d'achat"], dtype=float),
                        where=df_sim["Tarif d'achat"] != 0
                    )
        else:
            st.error(f"Type de simulation non reconnu : '{sim_type}'.")
            return None, 0.0

        total_adjusted_qty_all_months = pd.Series(0.0, index=df_sim.index)
        total_final_amount_all_months = pd.Series(0.0, index=df_sim.index)

        for month_name in selected_months:
            forecast_qty_col_name = f"Qté Prév. {month_name}"
            forecast_amount_col_name = f"Montant Prév. {month_name} (€)"
            
            base_q_series = base_monthly_forecast_qty_map.get(month_name, pd.Series(0.0, index=df_sim.index)) # Default to 0 if month somehow missing
            base_q_series = pd.to_numeric(base_q_series, errors='coerce').fillna(0)
            cond_series = df_sim["Conditionnement"] # Already int, >0
            
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
        
        df_sim["Qté Totale Prév. (Mois Sel.)"] = total_adjusted_qty_all_months
        df_sim["Montant Total Prév. (€) (Mois Sel.)"] = total_final_amount_all_months

        id_cols_display = ["Fournisseur", "Référence Article", "Désignation Article", "Conditionnement", "Tarif d'achat"]
        n1_sales_cols_display = sorted([f"Ventes N-1 {m}" for m in selected_months if f"Ventes N-1 {m}" in df_sim.columns])
        qty_forecast_cols_display = sorted([f"Qté Prév. {m}" for m in selected_months if f"Qté Prév. {m}" in df_sim.columns])
        amt_forecast_cols_display = sorted([f"Montant Prév. {m} (€)" for m in selected_months if f"Montant Prév. {m} (€)" in df_sim.columns])
        
        df_sim.rename(columns={ # Rename for UI brevity
            "Qté Totale Prév. (Mois Sel.)": "Qté Tot Prév (Mois Sel.)",
            "Montant Total Prév. (€) (Mois Sel.)": "Mnt Tot Prév (€) (Mois Sel.)"
        }, inplace=True)
        total_cols_display = [ # Use new names
            "Vts N-1 Tot (Mois Sel.)",
            "Qté Tot Prév (Mois Sel.)",
            "Mnt Tot Prév (€) (Mois Sel.)"
        ]

        final_ordered_cols = id_cols_display + total_cols_display + n1_sales_cols_display + qty_forecast_cols_display + amt_forecast_cols_display
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
    sanitized = re.sub(r'[\[\]:*?/\\<>|"]', '_', name) # Replace invalid chars
    if sanitized.startswith("'"): sanitized = "_" + sanitized[1:] # Avoid starting with '
    if sanitized.endswith("'"): sanitized = sanitized[:-1] + "_" # Avoid ending with '
    return sanitized[:31] # Truncate

def render_supplier_checkboxes(tab_key_prefix, all_suppliers, default_select_all=False):
    """ Renders supplier checkboxes with select/deselect all functionality for a tab. """
    select_all_key = f"{tab_key_prefix}_select_all"
    supplier_cb_keys = {
        supplier: f"{tab_key_prefix}_cb_{sanitize_supplier_key(supplier)}" for supplier in all_suppliers
    }

    # Initialize states if not present
    if select_all_key not in st.session_state:
        st.session_state[select_all_key] = default_select_all
        for cb_key in supplier_cb_keys.values():
            if cb_key not in st.session_state: # Only init if truly new
                st.session_state[cb_key] = default_select_all
    else: # select_all_key exists, ensure individual keys also exist
        for cb_key in supplier_cb_keys.values():
            if cb_key not in st.session_state:
                 st.session_state[cb_key] = st.session_state[select_all_key] # Default to current select_all state

    # Callbacks
    def toggle_all_suppliers_for_tab():
        current_select_all_value = st.session_state[select_all_key]
        logging.debug(f"Tab '{tab_key_prefix}': 'Select All' toggled to {current_select_all_value}.")
        for cb_key in supplier_cb_keys.values():
            st.session_state[cb_key] = current_select_all_value

    def check_individual_supplier_for_tab():
        all_individual_checked = all(
            st.session_state.get(cb_key, False) for cb_key in supplier_cb_keys.values()
        )
        if st.session_state.get(select_all_key) != all_individual_checked: # Check .get for safety
            st.session_state[select_all_key] = all_individual_checked
            # logging.debug(f"Tab '{tab_key_prefix}': 'Select All' auto-updated to {all_individual_checked}.")
    
    expander_label = "👤 Sélectionner Fournisseurs"
    if tab_key_prefix == "tab5":
        expander_label = "👤 Sélectionner Fournisseurs pour Export Suivi"

    with st.expander(expander_label, expanded=True):
        st.checkbox(
            "Sélectionner / Désélectionner Tout",
            key=select_all_key,
            on_change=toggle_all_suppliers_for_tab,
            disabled=not bool(all_suppliers)
        )
        st.markdown("---")

        selected_suppliers_in_ui = []
        num_display_cols = 4
        checkbox_cols = st.columns(num_display_cols)
        current_col_idx = 0
        
        for supplier_name, cb_key in supplier_cb_keys.items():
            # Checkbox state is implicitly taken from st.session_state[cb_key]
            checkbox_cols[current_col_idx].checkbox(
                supplier_name,
                key=cb_key, # This key links the widget to session_state
                on_change=check_individual_supplier_for_tab
            )
            if st.session_state.get(cb_key): # Read the current state
                selected_suppliers_in_ui.append(supplier_name)
            
            current_col_idx = (current_col_idx + 1) % num_display_cols
    return selected_suppliers_in_ui

def sanitize_supplier_key(supplier_name):
     """Creates a safe key for session state from supplier name."""
     if not isinstance(supplier_name, str):
         supplier_name = str(supplier_name)
     s = re.sub(r'\W+', '_', supplier_name) # Replace non-alphanumeric with _
     s = re.sub(r'^_+|_+$', '', s)          # Remove leading/trailing _
     s = re.sub(r'_+', '_', s)              # Consolidate multiple _
     return s if s else "invalid_supplier_key" # Handle empty string case

# --- Streamlit App Main Logic ---
st.set_page_config(page_title="Forecast & Rotation App", layout="wide")
st.title("📦 Application Prévision Commande, Analyse Rotation & Suivi")

# --- File Upload ---
uploaded_file = st.file_uploader(
    "📁 Charger le fichier Excel principal (contenant 'Tableau final', 'Minimum de commande', 'Suivi commandes')",
    type=["xlsx", "xls"],
    key="main_file_uploader"
)

# --- Initialize Session State ---
def get_default_session_state():
    return {
        'df_full': None,
        'min_order_dict': {},
        'df_initial_filtered': pd.DataFrame(), # Data after initial structural filters (Tableau Final based)
        'all_available_semaine_columns': [],
        'unique_suppliers_list': [], # Suppliers from df_initial_filtered

        # Tab 1 - Commande
        'commande_result_df': None,
        'commande_calculated_total_amount': 0.0,
        'commande_suppliers_calculated_for': [],

        # Tab 2 - Rotation
        'rotation_result_df': None,
        'rotation_analysis_period_label': "",
        'rotation_suppliers_calculated_for': [],
        'rotation_threshold_value': 1.0,
        'show_all_rotation_data': True,

        # Tab 4 - Forecast
        'forecast_result_df': None,
        'forecast_grand_total_amount': 0.0,
        'forecast_simulation_params_calculated_for': {},
        'forecast_selected_months_ui': list(calendar.month_name)[1:], # Default all months
        'forecast_sim_type_radio_index': 0, # Default to 'Simple Progression'
        'forecast_progression_percentage_ui': 5.0,
        'forecast_target_amount_ui': 10000.0,
        
        # Tab 5 - Suivi Commandes
        'df_suivi_commandes': None, # DataFrame from 'Suivi commandes' sheet
    }

# Initialize session state keys if they don't exist
for key, default_value in get_default_session_state().items():
    if key not in st.session_state:
        st.session_state[key] = default_value

# --- Data Loading and Initial Processing Block ---
if uploaded_file and st.session_state.df_full is None: # Process only if a new file is uploaded
    logging.info(f"New file uploaded: {uploaded_file.name}. Starting processing...")
    
    # --- Clear previous data and relevant UI states ---
    keys_to_reset_on_new_file = list(get_default_session_state().keys())
    dynamic_key_prefixes_to_clear = ['tab1_', 'tab2_', 'tab3_', 'tab4_', 'tab5_']

    for key in keys_to_reset_on_new_file:
        if key in st.session_state: del st.session_state[key]
    
    for prefix in dynamic_key_prefixes_to_clear:
        keys_to_remove = [k for k in st.session_state if k.startswith(prefix)]
        for k_to_remove in keys_to_remove: del st.session_state[k_to_remove]

    # Re-initialize with defaults after clearing
    for key, default_value in get_default_session_state().items():
        st.session_state[key] = default_value
    logging.info("Session state has been reset and re-initialized for the new file.")

    try:
        excel_file_buffer = io.BytesIO(uploaded_file.getvalue())
        
        # --- Read 'Tableau final' ---
        st.info("Lecture de l'onglet 'Tableau final'...")
        df_full_temp = safe_read_excel(excel_file_buffer, sheet_name="Tableau final", header=7)
        
        if df_full_temp is None:
            st.error("❌ Échec de la lecture de l'onglet 'Tableau final'. Vérifiez le nom et la structure.")
            st.stop()

        required_cols_tf = ["Stock", "Fournisseur", "AF_RefFourniss", "Tarif d'achat", "Conditionnement", "Référence Article", "Désignation Article"]
        missing_cols_tf_check = [col for col in required_cols_tf if col not in df_full_temp.columns]
        if missing_cols_tf_check:
            st.error(f"❌ Colonnes manquantes dans 'Tableau final': {', '.join(missing_cols_tf_check)}.")
            st.stop()

        df_full_temp["Stock"] = pd.to_numeric(df_full_temp["Stock"], errors='coerce').fillna(0)
        df_full_temp["Tarif d'achat"] = pd.to_numeric(df_full_temp["Tarif d'achat"], errors='coerce').fillna(0)
        df_full_temp["Conditionnement"] = pd.to_numeric(df_full_temp["Conditionnement"], errors='coerce').fillna(1).apply(lambda x: int(x) if x > 0 else 1)
        for str_col in ["Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article"]:
            if str_col in df_full_temp.columns: # Check existence before astype
                df_full_temp[str_col] = df_full_temp[str_col].astype(str).str.strip()
        st.session_state.df_full = df_full_temp
        st.success("✅ Onglet 'Tableau final' lu et traité.")

        # --- Read 'Minimum de commande' ---
        st.info("Lecture de l'onglet 'Minimum de commande'...")
        excel_file_buffer.seek(0) # Reset buffer pointer
        df_min_commande_temp = safe_read_excel(excel_file_buffer, sheet_name="Minimum de commande")
        min_order_dict_temp = {}
        if df_min_commande_temp is not None:
            supplier_col_min = "Fournisseur"
            min_amount_col = "Minimum de Commande"
            if supplier_col_min in df_min_commande_temp.columns and min_amount_col in df_min_commande_temp.columns:
                try:
                    df_min_commande_temp[supplier_col_min] = df_min_commande_temp[supplier_col_min].astype(str).str.strip()
                    df_min_commande_temp[min_amount_col] = pd.to_numeric(df_min_commande_temp[min_amount_col], errors='coerce')
                    min_order_dict_temp = df_min_commande_temp.dropna(subset=[supplier_col_min, min_amount_col]).set_index(supplier_col_min)[min_amount_col].to_dict()
                    st.success(f"✅ Onglet 'Minimum de commande' lu. {len(min_order_dict_temp)} minimums chargés.")
                except Exception as e_min_proc:
                    st.error(f"❌ Erreur traitement 'Minimum de commande': {e_min_proc}")
            else:
                st.warning(f"⚠️ Colonnes '{supplier_col_min}' et/ou '{min_amount_col}' non trouvées dans 'Minimum de commande'.")
        st.session_state.min_order_dict = min_order_dict_temp

        # --- Read 'Suivi commandes' ---
        st.info("Lecture onglet 'Suivi commandes'...")
        excel_file_buffer.seek(0) # Reset buffer pointer
        df_suivi_temp = safe_read_excel(excel_file_buffer, sheet_name="Suivi commandes", header=4) # header=4 for line 5
        
        if df_suivi_temp is not None:
            required_suivi_cols = ["Date Pièce BC", "N° de pièce", "AF_RefFourniss", "Désignation Article", "Qté Commandées", "Fournisseur"]
            missing_suivi_cols_check = [col for col in required_suivi_cols if col not in df_suivi_temp.columns]
            
            if not missing_suivi_cols_check:
                # Clean data if columns exist
                for col_to_strip in ["Fournisseur", "AF_RefFourniss", "Désignation Article", "N° de pièce"]:
                    if col_to_strip in df_suivi_temp.columns:
                         df_suivi_temp[col_to_strip] = df_suivi_temp[col_to_strip].astype(str).str.strip()
                if "Qté Commandées" in df_suivi_temp.columns:
                    df_suivi_temp["Qté Commandées"] = pd.to_numeric(df_suivi_temp["Qté Commandées"], errors='coerce').fillna(0)
                if "Date Pièce BC" in df_suivi_temp.columns:
                    try:
                        df_suivi_temp["Date Pièce BC"] = pd.to_datetime(df_suivi_temp["Date Pièce BC"], errors='coerce')
                    except Exception as e_date_parse_suivi:
                        st.warning(f"⚠️ Problème parsing 'Date Pièce BC' (Suivi commandes): {e_date_parse_suivi}.")
                
                # Drop rows where all values are NaN (often happens if header skips too many empty rows)
                df_suivi_temp.dropna(how='all', inplace=True)

                st.session_state.df_suivi_commandes = df_suivi_temp
                st.success(f"✅ Onglet 'Suivi commandes' lu ({len(df_suivi_temp)} lignes après nettoyage).")
            else:
                st.warning(f"⚠️ Colonnes manquantes dans 'Suivi commandes' (après lecture ligne 5): {', '.join(missing_suivi_cols_check)}. Suivi limité.")
                st.session_state.df_suivi_commandes = pd.DataFrame()
        else:
            st.info("Onglet 'Suivi commandes' non trouvé ou vide. Fonctionnalité de suivi non disponible.")
            st.session_state.df_suivi_commandes = pd.DataFrame()


        # --- Initial filtering from df_full for other tabs ---
        df_loaded_for_filter = st.session_state.df_full # Should be populated by now
        df_init_filtered_temp = df_loaded_for_filter[
            (df_loaded_for_filter["Fournisseur"].notna()) & (df_loaded_for_filter["Fournisseur"] != "") & (df_loaded_for_filter["Fournisseur"] != "#FILTER") &
            (df_loaded_for_filter["AF_RefFourniss"].notna()) & (df_loaded_for_filter["AF_RefFourniss"] != "")
        ].copy()
        st.session_state.df_initial_filtered = df_init_filtered_temp

        # Identify sales week columns
        first_potential_week_col_index = 12
        potential_sales_cols_list = []
        if len(df_loaded_for_filter.columns) > first_potential_week_col_index:
            candidate_cols_sales = df_loaded_for_filter.columns[first_potential_week_col_index:].tolist()
            known_non_week_cols_list = [
                "Tarif d'achat", "Conditionnement", "Stock", "Total", "Stock à terme", 
                "Ventes N-1", "Ventes 12 semaines identiques N-1", "Ventes 12 dernières semaines", 
                "Quantité à commander", "Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article"
            ] # Re-list for clarity for this specific operation
            exclude_from_sales_set = set(known_non_week_cols_list)
            for col_cand in candidate_cols_sales:
                if col_cand not in exclude_from_sales_set:
                    if pd.api.types.is_numeric_dtype(df_loaded_for_filter.get(col_cand, pd.Series(dtype=object)).dtype):
                        potential_sales_cols_list.append(col_cand)
        st.session_state.all_available_semaine_columns = potential_sales_cols_list
        if not potential_sales_cols_list:
            st.warning("⚠️ Aucune colonne de vente numérique n'a été automatiquement identifiée.")
        
        # Populate unique suppliers list from the df_initial_filtered data (for tabs 1, 2, 4)
        if not df_init_filtered_temp.empty and "Fournisseur" in df_init_filtered_temp.columns:
            st.session_state.unique_suppliers_list = sorted(df_init_filtered_temp["Fournisseur"].unique().tolist())
        
        st.rerun() # Rerun to update UI with loaded data

    except Exception as e_load_main_block:
        st.error(f"❌ Une erreur majeure est survenue lors du chargement ou du traitement initial du fichier : {e_load_main_block}")
        logging.exception("Major file loading/processing error in main block:")
        st.session_state.df_full = None # Allow re-upload attempt
        st.stop()

# --- Main Application UI (Tabs) ---
if 'df_initial_filtered' in st.session_state and isinstance(st.session_state.df_initial_filtered, pd.DataFrame):

    df_base_for_tabs = st.session_state.df_initial_filtered
    all_suppliers_from_data = st.session_state.unique_suppliers_list # From Tableau Final (filtered)
    min_order_amounts = st.session_state.min_order_dict
    identified_semaine_cols = st.session_state.all_available_semaine_columns
    df_suivi_commandes_all_data = st.session_state.get('df_suivi_commandes', pd.DataFrame())

    tab_titles = ["Prévision Commande", "Analyse Rotation Stock", "Vérification Stock", "Simulation Forecast", "Suivi Commandes Fourn."]
    tab1, tab2, tab3, tab4, tab5 = st.tabs(tab_titles)

    # ========================= TAB 1: Prévision Commande =========================
    with tab1:
        st.header("Prévision des Quantités à Commander")
        selected_fournisseurs_tab1 = render_supplier_checkboxes("tab1", all_suppliers_from_data, default_select_all=True)
        
        df_display_tab1 = pd.DataFrame() 
        if selected_fournisseurs_tab1:
            if not df_base_for_tabs.empty:
                df_display_tab1 = df_base_for_tabs[df_base_for_tabs["Fournisseur"].isin(selected_fournisseurs_tab1)].copy()
                st.caption(f"{len(df_display_tab1)} articles pour {len(selected_fournisseurs_tab1)} fournisseur(s) sélectionné(s).")
            else: st.caption("Aucune donnée de base à filtrer pour les fournisseurs.")
        else: st.info("Veuillez sélectionner au moins un fournisseur.")
        st.markdown("---")

        if df_display_tab1.empty and selected_fournisseurs_tab1 :
            st.warning("Aucun article trouvé pour le(s) fournisseur(s) sélectionné(s) dans les données de base.")
        elif not identified_semaine_cols and not df_display_tab1.empty:
            st.warning("Impossible de calculer : Aucune colonne de ventes (semaines) n'a été identifiée dans le fichier.")
        elif not df_display_tab1.empty :
            st.markdown("#### Paramètres de Calcul de Commande")
            col1_params_cmd, col2_params_cmd = st.columns(2)
            with col1_params_cmd:
                duree_couverture_semaines_cmd_ui = st.number_input(
                    label="⏳ Durée de couverture souhaitée (en semaines)", 
                    min_value=1, max_value=260, value=4, step=1, key="duree_couv_cmd_ui_tab1"
                )
            with col2_params_cmd:
                montant_min_global_cmd_ui = st.number_input(
                    label="💶 Montant minimum global de commande (€)", 
                    min_value=0.0, value=0.0, step=50.0, format="%.2f", key="montant_min_cmd_ui_tab1"
                )
            
            if st.button("🚀 Calculer les Quantités à Commander", key="calc_qte_cmd_btn_main_tab1"):
                with st.spinner("Calcul des quantités en cours..."):
                    result_tuple_cmd_calc = calculer_quantite_a_commander(
                        df_display_tab1, 
                        identified_semaine_cols, 
                        montant_min_global_cmd_ui, 
                        duree_couverture_semaines_cmd_ui
                    )
                
                if result_tuple_cmd_calc:
                    st.success("✅ Calcul des quantités terminé.")
                    quantites_calculees_res, ventes_n1_calc_res, ventes_12_n1_sim_calc_res, ventes_12_dern_calc_res, montant_total_calc_res = result_tuple_cmd_calc
                    
                    df_result_commande = df_display_tab1.copy()
                    df_result_commande["Qte Cmdée"] = quantites_calculees_res
                    df_result_commande["Vts N-1 Total (calc)"] = ventes_n1_calc_res
                    df_result_commande["Vts 12 N-1 Sim (calc)"] = ventes_12_n1_sim_calc_res
                    df_result_commande["Vts 12 Dern. (calc)"] = ventes_12_dern_calc_res
                    df_result_commande["Tarif Ach."] = pd.to_numeric(df_result_commande["Tarif d'achat"], errors='coerce').fillna(0)
                    df_result_commande["Total Cmd (€)"] = df_result_commande["Tarif Ach."] * df_result_commande["Qte Cmdée"]
                    df_result_commande["Stock Terme"] = df_result_commande["Stock"] + df_result_commande["Qte Cmdée"]
                    
                    st.session_state.commande_result_df = df_result_commande
                    st.session_state.commande_calculated_total_amount = montant_total_calc_res
                    st.session_state.commande_suppliers_calculated_for = selected_fournisseurs_tab1
                    st.rerun()
                else:
                    st.error("❌ Le calcul des quantités a échoué ou n'a retourné aucun résultat.")

            if st.session_state.commande_result_df is not None:
                if st.session_state.commande_suppliers_calculated_for == selected_fournisseurs_tab1:
                    st.markdown("---")
                    st.markdown("#### Résultats de la Prévision de Commande")
                    
                    df_to_display_commande_final = st.session_state.commande_result_df
                    calculated_total_commande_final = st.session_state.commande_calculated_total_amount
                    suppliers_in_calc_commande_final = st.session_state.commande_suppliers_calculated_for

                    st.metric(label="💰 Montant Total Commandé (calculé)", value=f"{calculated_total_commande_final:,.2f} €")

                    if len(suppliers_in_calc_commande_final) == 1:
                        single_supplier_name_cmd = suppliers_in_calc_commande_final[0]
                        if single_supplier_name_cmd in min_order_amounts:
                            required_min_for_supplier_cmd = min_order_amounts[single_supplier_name_cmd]
                            actual_total_for_supplier_cmd = df_to_display_commande_final[
                                df_to_display_commande_final["Fournisseur"] == single_supplier_name_cmd
                            ]["Total Cmd (€)"].sum()
                            if required_min_for_supplier_cmd > 0 and actual_total_for_supplier_cmd < required_min_for_supplier_cmd:
                                difference_cmd = required_min_for_supplier_cmd - actual_total_for_supplier_cmd
                                st.warning(f"⚠️ Min non atteint ({single_supplier_name_cmd}): {actual_total_for_supplier_cmd:,.2f}€ / Requis: {required_min_for_supplier_cmd:,.2f}€ (Manque: {difference_cmd:,.2f}€)")
                    
                    cols_to_show_commande_final = ["Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article", "Stock", "Vts N-1 Total (calc)", "Vts 12 N-1 Sim (calc)", "Vts 12 Dern. (calc)", "Conditionnement", "Qte Cmdée", "Stock Terme", "Tarif Ach.", "Total Cmd (€)"]
                    displayable_cols_commande_final = [col for col in cols_to_show_commande_final if col in df_to_display_commande_final.columns]
                    
                    if not displayable_cols_commande_final: st.error("Aucune colonne de résultat à afficher (commande).")
                    else:
                        formatters_commande_final = {"Tarif Ach.": "{:,.2f}€", "Total Cmd (€)": "{:,.2f}€", "Vts N-1 Total (calc)": "{:,.0f}", "Vts 12 N-1 Sim (calc)": "{:,.0f}", "Vts 12 Dern. (calc)": "{:,.0f}", "Stock": "{:,.0f}", "Conditionnement": "{:,.0f}", "Qte Cmdée": "{:,.0f}", "Stock Terme": "{:,.0f}"}
                        st.dataframe(df_to_display_commande_final[displayable_cols_commande_final].style.format(formatters_commande_final, na_rep="-", thousands=","))

                    st.markdown("#### Exporter les Commandes Calculées")
                    df_export_commande_base = df_to_display_commande_final[df_to_display_commande_final["Qte Cmdée"] > 0].copy()
                    if not df_export_commande_base.empty:
                        excel_output_buffer_commande = io.BytesIO()
                        sheets_created_commande = 0
                        try:
                            with pd.ExcelWriter(excel_output_buffer_commande, engine="openpyxl") as writer_commande_export:
                                export_cols_per_sheet_cmd_exp = [c for c in displayable_cols_commande_final if c != 'Fournisseur']
                                qty_col_name_exp = "Qte Cmdée"; price_col_name_exp = "Tarif Ach."; total_col_name_exp = "Total Cmd (€)"
                                formula_possible_exp = False
                                if all(c in export_cols_per_sheet_cmd_exp for c in [qty_col_name_exp, price_col_name_exp, total_col_name_exp]):
                                    try:
                                        qty_col_letter_exp = get_column_letter(export_cols_per_sheet_cmd_exp.index(qty_col_name_exp) + 1)
                                        price_col_letter_exp = get_column_letter(export_cols_per_sheet_cmd_exp.index(price_col_name_exp) + 1)
                                        total_col_letter_exp = get_column_letter(export_cols_per_sheet_cmd_exp.index(total_col_name_exp) + 1)
                                        formula_possible_exp = True
                                    except ValueError: pass

                                for supplier_for_sheet_exp in suppliers_in_calc_commande_final:
                                    df_supplier_sheet_data_exp = df_export_commande_base[df_export_commande_base["Fournisseur"] == supplier_for_sheet_exp]
                                    if not df_supplier_sheet_data_exp.empty:
                                        df_to_write_to_sheet_exp = df_supplier_sheet_data_exp[export_cols_per_sheet_cmd_exp].copy()
                                        num_data_rows_exp = len(df_to_write_to_sheet_exp)
                                        label_col_for_summary_exp = "Désignation Article" if "Désignation Article" in export_cols_per_sheet_cmd_exp else (export_cols_per_sheet_cmd_exp[1] if len(export_cols_per_sheet_cmd_exp) > 1 else export_cols_per_sheet_cmd_exp[0])
                                        total_value_for_sheet_exp = df_to_write_to_sheet_exp[total_col_name_exp].sum()
                                        min_req_for_sheet_exp = min_order_amounts.get(supplier_for_sheet_exp, 0)
                                        min_req_display_exp = f"{min_req_for_sheet_exp:,.2f}€" if min_req_for_sheet_exp > 0 else "N/A"
                                        summary_rows_data_exp = [{label_col_for_summary_exp: "TOTAL", total_col_name_exp: total_value_for_sheet_exp}, {label_col_for_summary_exp: "Minimum Requis Fournisseur", total_col_name_exp: min_req_display_exp}]
                                        df_summary_rows_exp = pd.DataFrame(summary_rows_data_exp, columns=export_cols_per_sheet_cmd_exp).fillna('')
                                        df_final_sheet_exp = pd.concat([df_to_write_to_sheet_exp, df_summary_rows_exp], ignore_index=True)
                                        safe_sheet_name_exp = sanitize_sheet_name(supplier_for_sheet_exp)
                                        try:
                                            df_final_sheet_exp.to_excel(writer_commande_export, sheet_name=safe_sheet_name_exp, index=False)
                                            worksheet_exp = writer_commande_export.sheets[safe_sheet_name_exp]
                                            if formula_possible_exp and num_data_rows_exp > 0:
                                                for r_idx_exp in range(2, num_data_rows_exp + 2): # Excel rows are 1-based, data starts at row 2
                                                    worksheet_exp[f"{total_col_letter_exp}{r_idx_exp}"].value = f"={qty_col_letter_exp}{r_idx_exp}*{price_col_letter_exp}{r_idx_exp}"
                                                    worksheet_exp[f"{total_col_letter_exp}{r_idx_exp}"].number_format = '#,##0.00€'
                                                total_formula_cell_loc_exp = f"{total_col_letter_exp}{num_data_rows_exp + 2}" # TOTAL row
                                                worksheet_exp[total_formula_cell_loc_exp].value = f"=SUM({total_col_letter_exp}2:{total_col_letter_exp}{num_data_rows_exp + 1})"
                                                worksheet_exp[total_formula_cell_loc_exp].number_format = '#,##0.00€'
                                            sheets_created_commande += 1
                                        except Exception as e_sheet_write_exp: logging.error(f"Erreur écriture feuille Excel '{safe_sheet_name_exp}': {e_sheet_write_exp}")
                            if sheets_created_commande > 0:
                                writer_commande_export.save()
                                excel_output_buffer_commande.seek(0)
                                export_filename_commande = f"commandes_{'multiples_fournisseurs' if len(suppliers_in_calc_commande_final) > 1 else sanitize_sheet_name(suppliers_in_calc_commande_final[0])}_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx"
                                st.download_button(label=f"📥 Télécharger Commandes ({sheets_created_commande} feuille(s))", data=excel_output_buffer_commande, file_name=export_filename_commande, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_cmd_export_btn_main_tab1")
                            else: st.info("Aucune donnée de commande à exporter (quantités commandées pourraient être nulles).")
                        except Exception as e_excel_writer_cmd: logging.exception(f"Erreur ExcelWriter pour commandes: {e_excel_writer_cmd}"); st.error("Une erreur est survenue lors de la préparation du fichier Excel pour les commandes.")
                    else: st.info("Aucun article avec une quantité commandée > 0 à exporter.")
                else: st.info("Les résultats de commande affichés précédemment sont invalidés car la sélection de fournisseurs a changé. Veuillez relancer le calcul.")

    # ====================== TAB 2: Analyse Rotation Stock ======================
    with tab2:
        st.header("Analyse de la Rotation des Stocks")
        selected_fournisseurs_tab2 = render_supplier_checkboxes("tab2", all_suppliers_from_data, default_select_all=True)
        
        df_display_tab2 = pd.DataFrame()
        if selected_fournisseurs_tab2:
            if not df_base_for_tabs.empty:
                df_display_tab2 = df_base_for_tabs[df_base_for_tabs["Fournisseur"].isin(selected_fournisseurs_tab2)].copy()
                st.caption(f"{len(df_display_tab2)} articles pour {len(selected_fournisseurs_tab2)} fournisseur(s) sélectionné(s).")
            else: st.caption("Aucune donnée de base à filtrer.")
        else: st.info("Veuillez sélectionner au moins un fournisseur.")
        st.markdown("---")

        if df_display_tab2.empty and selected_fournisseurs_tab2:
            st.warning("Aucun article trouvé pour le(s) fournisseur(s) sélectionné(s).")
        elif not identified_semaine_cols and not df_display_tab2.empty:
            st.warning("Impossible d'analyser : Aucune colonne de ventes (semaines) n'a été identifiée.")
        elif not df_display_tab2.empty:
            st.markdown("#### Paramètres d'Analyse de Rotation")
            col1_rot_params_ui, col2_rot_params_ui = st.columns(2)
            with col1_rot_params_ui:
                period_options_rot_ui = {"12 dernières semaines": 12, "52 dernières semaines": 52, "Toutes les données disponibles": 0}
                selected_period_label_rot_ui = st.selectbox("⏳ Période d'analyse des ventes:", options=period_options_rot_ui.keys(), key="rot_analysis_period_selectbox_tab2")
                selected_period_weeks_rot_ui = period_options_rot_ui[selected_period_label_rot_ui]
            with col2_rot_params_ui:
                st.markdown("##### Options d'Affichage des Résultats")
                show_all_rot_ui_val = st.checkbox("Afficher tous les articles", value=st.session_state.show_all_rotation_data, key="show_all_rot_ui_cb_main_tab2")
                st.session_state.show_all_rotation_data = show_all_rot_ui_val
                rotation_filter_threshold_ui_val = st.number_input("... ou afficher articles avec ventes mensuelles moyennes <", min_value=0.0, value=st.session_state.rotation_threshold_value, step=0.1, format="%.1f", key="rot_filter_threshold_ui_numin_main_tab2", disabled=show_all_rot_ui_val)
                if not show_all_rot_ui_val: st.session_state.rotation_threshold_value = rotation_filter_threshold_ui_val

            if st.button("🔄 Analyser la Rotation des Stocks", key="analyze_rotation_btn_main_tab2"):
                with st.spinner("Analyse de la rotation en cours..."):
                    df_rotation_results_calc = calculer_rotation_stock(df_display_tab2, identified_semaine_cols, selected_period_weeks_rot_ui)
                if df_rotation_results_calc is not None:
                    st.success("✅ Analyse de rotation terminée.")
                    st.session_state.rotation_result_df = df_rotation_results_calc
                    st.session_state.rotation_analysis_period_label = selected_period_label_rot_ui
                    st.session_state.rotation_suppliers_calculated_for = selected_fournisseurs_tab2
                    st.rerun()
                else: st.error("❌ L'analyse de rotation a échoué ou n'a pas produit de résultats.")
            
            if st.session_state.rotation_result_df is not None:
                if st.session_state.rotation_suppliers_calculated_for == selected_fournisseurs_tab2:
                    st.markdown("---")
                    st.markdown(f"#### Résultats de l'Analyse de Rotation ({st.session_state.rotation_analysis_period_label})")
                    df_rotation_output_base_res = st.session_state.rotation_result_df
                    current_filter_threshold_res = st.session_state.rotation_threshold_value
                    show_all_filter_active_res = st.session_state.show_all_rotation_data
                    df_filtered_for_display_rot_res = pd.DataFrame()
                    monthly_sales_col_name_rot_res = "Ventes Moy Mensuel (Période)"

                    if df_rotation_output_base_res.empty: st.info("Aucune donnée de rotation à afficher.")
                    elif show_all_filter_active_res:
                        df_filtered_for_display_rot_res = df_rotation_output_base_res.copy()
                        st.caption(f"Affichage de {len(df_filtered_for_display_rot_res)} articles (tous).")
                    elif monthly_sales_col_name_rot_res in df_rotation_output_base_res.columns:
                        try:
                            sales_for_filter_res = pd.to_numeric(df_rotation_output_base_res[monthly_sales_col_name_rot_res], errors='coerce').fillna(0)
                            df_filtered_for_display_rot_res = df_rotation_output_base_res[sales_for_filter_res < current_filter_threshold_res].copy()
                            st.caption(f"Filtré : Vts < {current_filter_threshold_res:.1f}/mois. {len(df_filtered_for_display_rot_res)} / {len(df_rotation_output_base_res)} art.")
                            if df_filtered_for_display_rot_res.empty: st.info(f"Aucun article < {current_filter_threshold_res:.1f} vts/mois.")
                        except Exception as e_filter_rot_res: st.error(f"Err filtre rotation: {e_filter_rot_res}"); df_filtered_for_display_rot_res = df_rotation_output_base_res.copy()
                    else:
                        st.warning(f"Col '{monthly_sales_col_name_rot_res}' non trouvée. Affichage tout."); df_filtered_for_display_rot_res = df_rotation_output_base_res.copy()

                    if not df_filtered_for_display_rot_res.empty:
                        cols_to_show_rot_res = ["Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article", "Tarif d'achat", "Stock", "Unités Vendues (Période)", "Ventes Moy Hebdo (Période)", "Ventes Moy Mensuel (Période)", "Semaines Stock (WoS)", "Rotation Unités (Proxy)", "Valeur Stock Actuel (€)", "COGS (Période)", "Rotation Valeur (Proxy)"]
                        displayable_cols_rot_res = [col for col in cols_to_show_rot_res if col in df_filtered_for_display_rot_res.columns]
                        df_display_copy_rot_res = df_filtered_for_display_rot_res[displayable_cols_rot_res].copy()
                        numeric_cols_rounding_map_rot_res = {"Tarif d'achat": 2, "Ventes Moy Hebdo (Période)": 2, "Ventes Moy Mensuel (Période)": 2, "Semaines Stock (WoS)": 1, "Rotation Unités (Proxy)": 2, "Valeur Stock Actuel (€)": 2, "COGS (Période)": 2, "Rotation Valeur (Proxy)": 2}
                        for num_col_res, round_digits_res in numeric_cols_rounding_map_rot_res.items():
                            if num_col_res in df_display_copy_rot_res.columns:
                                df_display_copy_rot_res[num_col_res] = pd.to_numeric(df_display_copy_rot_res[num_col_res], errors='coerce').round(round_digits_res)
                        df_display_copy_rot_res.replace([np.inf, -np.inf], 'Infini', inplace=True)
                        formatters_rot_res = {"Tarif d'achat": "{:,.2f}€", "Stock": "{:,.0f}", "Unités Vendues (Période)": "{:,.0f}", "Ventes Moy Hebdo (Période)": "{:,.2f}", "Ventes Moy Mensuel (Période)": "{:,.2f}", "Semaines Stock (WoS)": "{}", "Rotation Unités (Proxy)": "{}", "Valeur Stock Actuel (€)": "{:,.2f}€", "COGS (Période)": "{:,.2f}€", "Rotation Valeur (Proxy)": "{}"}
                        st.dataframe(df_display_copy_rot_res.style.format(formatters_rot_res, na_rep="-", thousands=","))

                        st.markdown("#### Exporter l'Analyse de Rotation Affichée")
                        excel_output_buffer_rot_exp = io.BytesIO()
                        df_export_rot_final = df_display_copy_rot_res
                        sheet_name_label_rot_exp = f"Rotation_{'Filtree' if not show_all_filter_active_res else 'Complete'}"
                        export_filename_base_rot_exp = f"analyse_rotation_{'filtree' if not show_all_filter_active_res else 'complete'}"
                        current_suppliers_for_export_name_rot = 'multiples_fournisseurs' if len(selected_fournisseurs_tab2) > 1 else (sanitize_sheet_name(selected_fournisseurs_tab2[0]) if selected_fournisseurs_tab2 else 'aucun_fournisseur')
                        with pd.ExcelWriter(excel_output_buffer_rot_exp, engine="openpyxl") as writer_rot_exp:
                            df_export_rot_final.to_excel(writer_rot_exp, sheet_name=sanitize_sheet_name(sheet_name_label_rot_exp), index=False)
                        excel_output_buffer_rot_exp.seek(0)
                        export_filename_rot_final = f"{export_filename_base_rot_exp}_{current_suppliers_for_export_name_rot}_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx"
                        download_label_rot_final = f"📥 Télécharger Analyse {'Filtrée' if not show_all_filter_active_res else 'Complète'}"
                        st.download_button(label=download_label_rot_final, data=excel_output_buffer_rot_exp, file_name=export_filename_rot_final, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_rotation_export_btn_main_tab2")
                else: st.info("Les résultats d'analyse de rotation sont invalidés (sélection fournisseurs a changé). Relancer.")

    # ========================= TAB 3: Vérification Stock =========================
    with tab3:
        st.header("Vérification des Stocks Négatifs")
        st.caption("Cette analyse porte sur l'ensemble des articles du fichier 'Tableau final'.")
        df_full_for_neg_check_tab3 = st.session_state.get('df_full', None)

        if df_full_for_neg_check_tab3 is None or not isinstance(df_full_for_neg_check_tab3, pd.DataFrame):
            st.warning("Les données ('Tableau final') n'ont pas été chargées.")
        elif df_full_for_neg_check_tab3.empty:
            st.info("Le 'Tableau final' est vide.")
        else:
            stock_col_name_neg = "Stock"
            if stock_col_name_neg not in df_full_for_neg_check_tab3.columns:
                st.error(f"La colonne '{stock_col_name_neg}' est introuvable dans 'Tableau final'.")
            else:
                df_negative_stocks_res = df_full_for_neg_check_tab3[df_full_for_neg_check_tab3[stock_col_name_neg] < 0].copy()
                if df_negative_stocks_res.empty:
                    st.success("✅ Aucun article avec un stock négatif trouvé.")
                else:
                    st.warning(f"⚠️ **{len(df_negative_stocks_res)} article(s) trouvé(s) avec un stock négatif !**")
                    cols_to_show_neg_stock_tab3 = ["Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article", "Stock"]
                    displayable_cols_neg_stock_tab3 = [col for col in cols_to_show_neg_stock_tab3 if col in df_negative_stocks_res.columns]
                    if not displayable_cols_neg_stock_tab3: st.error("Impossible d'afficher détails stocks négatifs.")
                    else:
                        st.dataframe(df_negative_stocks_res[displayable_cols_neg_stock_tab3].style.format({"Stock": "{:,.0f}"}, na_rep="-").apply(lambda s: ['background-color:#FADBD8' if s.name == stock_col_name_neg and val < 0 else '' for val in s], axis=0)) # Check axis for apply
                        st.markdown("---")
                        st.markdown("#### Exporter la Liste des Stocks Négatifs")
                        excel_output_buffer_neg_tab3 = io.BytesIO()
                        df_export_neg_stock_tab3 = df_negative_stocks_res[displayable_cols_neg_stock_tab3].copy()
                        try:
                            with pd.ExcelWriter(excel_output_buffer_neg_tab3, engine="openpyxl") as writer_neg_tab3:
                                df_export_neg_stock_tab3.to_excel(writer_neg_tab3, sheet_name="Stocks_Negatifs", index=False)
                            excel_output_buffer_neg_tab3.seek(0)
                            export_filename_neg_tab3 = f"stocks_negatifs_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx"
                            st.download_button("📥 Télécharger la Liste des Stocks Négatifs", data=excel_output_buffer_neg_tab3, file_name=export_filename_neg_tab3, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_neg_stock_export_btn_main_tab3")
                        except Exception as e_export_neg_tab3: st.error(f"Erreur export stocks négatifs: {e_export_neg_tab3}")

    # ========================= TAB 4: Simulation Forecast =========================
    with tab4:
        st.header("Simulation de Forecast Annuel")
        selected_fournisseurs_tab4 = render_supplier_checkboxes("tab4", all_suppliers_from_data, default_select_all=True)
        
        df_display_tab4 = pd.DataFrame()
        if selected_fournisseurs_tab4:
            if not df_base_for_tabs.empty:
                df_display_tab4 = df_base_for_tabs[df_base_for_tabs["Fournisseur"].isin(selected_fournisseurs_tab4)].copy()
                st.caption(f"{len(df_display_tab4)} articles pour {len(selected_fournisseurs_tab4)} fournisseur(s) sélectionné(s).")
            else: st.caption("Aucune donnée de base à filtrer.")
        else: st.info("Veuillez sélectionner au moins un fournisseur.")
        st.markdown("---")
        st.warning("🚨 **Hypothèse importante :** Saisonnalité mensuelle basée sur découpage approx. des 52 sem. N-1.")

        if df_display_tab4.empty and selected_fournisseurs_tab4:
            st.warning("Aucun article trouvé pour le(s) fournisseur(s) sélectionné(s).")
        elif len(identified_semaine_cols) < 52 and not df_display_tab4.empty :
            st.warning(f"Données historiques insuffisantes ({len(identified_semaine_cols)} semaines trouvées). Simulation N-1 impossible.")
        elif not df_display_tab4.empty:
            st.markdown("#### Paramètres de Simulation de Forecast")
            all_calendar_months_ui = list(calendar.month_name)[1:]
            selected_months_for_forecast_ui_tab4 = st.multiselect("📅 Mois à inclure dans la simulation:", options=all_calendar_months_ui, default=st.session_state.forecast_selected_months_ui, key="fcst_months_multiselect_ui_main_tab4")
            st.session_state.forecast_selected_months_ui = selected_months_for_forecast_ui_tab4
            
            sim_type_options_fcst_ui = ('Simple Progression', 'Objectif Montant')
            selected_sim_type_fcst_ui_tab4 = st.radio("⚙️ Type de Simulation:", options=sim_type_options_fcst_ui, horizontal=True, index=st.session_state.forecast_sim_type_radio_index, key="fcst_sim_type_radio_ui_main_tab4")
            st.session_state.forecast_sim_type_radio_index = sim_type_options_fcst_ui.index(selected_sim_type_fcst_ui_tab4)
            
            prog_percentage_fcst_ui_tab4 = 0.0
            target_amount_fcst_ui_tab4 = 0.0
            col1_fcst_params_tab4, col2_fcst_params_tab4 = st.columns(2)
            with col1_fcst_params_tab4:
                if selected_sim_type_fcst_ui_tab4 == 'Simple Progression':
                    prog_percentage_fcst_ui_tab4 = st.number_input(label="📈 Taux de Progression (%) vs N-1", min_value=-100.0, value=st.session_state.forecast_progression_percentage_ui, step=0.5, format="%.1f", key="fcst_prog_pct_ui_numin_main_tab4")
                    st.session_state.forecast_progression_percentage_ui = prog_percentage_fcst_ui_tab4
            with col2_fcst_params_tab4:
                if selected_sim_type_fcst_ui_tab4 == 'Objectif Montant':
                    target_amount_fcst_ui_tab4 = st.number_input(label="🎯 Montant Objectif (€) (période sel.)", min_value=0.0, value=st.session_state.forecast_target_amount_ui, step=1000.0, format="%.2f", key="fcst_target_amt_ui_numin_main_tab4")
                    st.session_state.forecast_target_amount_ui = target_amount_fcst_ui_tab4

            if st.button("▶️ Lancer la Simulation de Forecast", key="run_forecast_sim_btn_main_tab4"):
                if not selected_months_for_forecast_ui_tab4: st.error("Veuillez sélectionner au moins un mois.")
                else:
                    with st.spinner("Simulation du forecast en cours..."):
                        df_forecast_sim_result_calc, grand_total_sim_amount_calc = calculer_forecast_simulation_v3(df_display_tab4, identified_semaine_cols, selected_months_for_forecast_ui_tab4, selected_sim_type_fcst_ui_tab4, prog_percentage_fcst_ui_tab4, target_amount_fcst_ui_tab4)
                    if df_forecast_sim_result_calc is not None:
                        st.success("✅ Simulation de forecast terminée.")
                        st.session_state.forecast_result_df = df_forecast_sim_result_calc
                        st.session_state.forecast_grand_total_amount = grand_total_sim_amount_calc
                        st.session_state.forecast_simulation_params_calculated_for = {'suppliers': selected_fournisseurs_tab4, 'months': selected_months_for_forecast_ui_tab4, 'type': selected_sim_type_fcst_ui_tab4, 'prog_pct': prog_percentage_fcst_ui_tab4 if selected_sim_type_fcst_ui_tab4 == 'Simple Progression' else 0.0, 'obj_amt': target_amount_fcst_ui_tab4 if selected_sim_type_fcst_ui_tab4 == 'Objectif Montant' else 0.0}
                        st.rerun()
                    else: st.error("❌ La simulation de forecast a échoué.")
            
            if st.session_state.forecast_result_df is not None:
                current_ui_params_fcst_tab4 = {'suppliers': selected_fournisseurs_tab4, 'months': selected_months_for_forecast_ui_tab4, 'type': selected_sim_type_fcst_ui_tab4, 'prog_pct': st.session_state.forecast_progression_percentage_ui if selected_sim_type_fcst_ui_tab4=='Simple Progression' else 0.0, 'obj_amt': st.session_state.forecast_target_amount_ui if selected_sim_type_fcst_ui_tab4=='Objectif Montant' else 0.0}
                if st.session_state.forecast_simulation_params_calculated_for == current_ui_params_fcst_tab4:
                    st.markdown("---")
                    st.markdown("#### Résultats de la Simulation de Forecast")
                    df_to_display_fcst_final = st.session_state.forecast_result_df
                    grand_total_fcst_disp_final = st.session_state.forecast_grand_total_amount
                    if df_to_display_fcst_final.empty: st.info("Aucun résultat de simulation à afficher.")
                    else:
                        formatters_fcst_final = {"Tarif d'achat": "{:,.2f}€", "Conditionnement": "{:,.0f}"}
                        for month_disp_fcst in selected_months_for_forecast_ui_tab4:
                            if f"Ventes N-1 {month_disp_fcst}" in df_to_display_fcst_final.columns: formatters_fcst_final[f"Ventes N-1 {month_disp_fcst}"] = "{:,.0f}"
                            if f"Qté Prév. {month_disp_fcst}" in df_to_display_fcst_final.columns: formatters_fcst_final[f"Qté Prév. {month_disp_fcst}"] = "{:,.0f}"
                            if f"Montant Prév. {month_disp_fcst} (€)" in df_to_display_fcst_final.columns: formatters_fcst_final[f"Montant Prév. {month_disp_fcst} (€)"] = "{:,.2f}€"
                        if "Vts N-1 Tot (Mois Sel.)" in df_to_display_fcst_final.columns: formatters_fcst_final["Vts N-1 Tot (Mois Sel.)"] = "{:,.0f}"
                        if "Qté Tot Prév (Mois Sel.)" in df_to_display_fcst_final.columns: formatters_fcst_final["Qté Tot Prév (Mois Sel.)"] = "{:,.0f}"
                        if "Mnt Tot Prév (€) (Mois Sel.)" in df_to_display_fcst_final.columns: formatters_fcst_final["Mnt Tot Prév (€) (Mois Sel.)"] = "{:,.2f}€"
                        try: st.dataframe(df_to_display_fcst_final.style.format(formatters_fcst_final, na_rep="-", thousands=","))
                        except Exception as e_fmt_fcst_final: st.error(f"Erreur formatage affichage forecast: {e_fmt_fcst_final}"); st.dataframe(df_to_display_fcst_final)
                        st.metric(label="💰 Montant Total Général Prévisionnel (€) (mois sélectionnés)", value=f"{grand_total_fcst_disp_final:,.2f} €")

                        st.markdown("#### Exporter la Simulation de Forecast")
                        excel_output_buffer_fcst_exp = io.BytesIO()
                        df_export_fcst_final = df_to_display_fcst_final.copy()
                        try:
                            sim_type_filename_part_exp = selected_sim_type_fcst_ui_tab4.replace(' ', '_').lower()
                            with pd.ExcelWriter(excel_output_buffer_fcst_exp, engine="openpyxl") as writer_fcst_exp:
                                df_export_fcst_final.to_excel(writer_fcst_exp, sheet_name=sanitize_sheet_name(f"Forecast_{sim_type_filename_part_exp}"), index=False)
                            excel_output_buffer_fcst_exp.seek(0)
                            current_suppliers_for_export_name_fcst_exp = 'multiples_fournisseurs' if len(selected_fournisseurs_tab4) > 1 else (sanitize_sheet_name(selected_fournisseurs_tab4[0]) if selected_fournisseurs_tab4 else 'aucun_fournisseur')
                            export_filename_fcst_final = f"simulation_forecast_{sim_type_filename_part_exp}_{current_suppliers_for_export_name_fcst_exp}_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx"
                            st.download_button("📥 Télécharger la Simulation", data=excel_output_buffer_fcst_exp, file_name=export_filename_fcst_final, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_forecast_sim_export_btn_main_tab4")
                        except Exception as e_export_fcst_final: st.error(f"Erreur export simulation: {e_export_fcst_final}")
                else: st.info("Les résultats de simulation sont invalidés (paramètres/fournisseurs ont changé). Relancer.")

    # ========================= TAB 5: Suivi Commandes Fournisseurs =========================
    with tab5:
        st.header("📄 Suivi des Commandes Fournisseurs")

        if df_suivi_commandes_all_data is None or df_suivi_commandes_all_data.empty:
            st.warning("Aucune donnée de suivi de commandes n'a été chargée (onglet 'Suivi commandes' vide/manquant ou erreur de lecture).")
        else:
            suppliers_in_suivi_list = []
            if "Fournisseur" in df_suivi_commandes_all_data.columns:
                suppliers_in_suivi_list = sorted(df_suivi_commandes_all_data["Fournisseur"].astype(str).unique().tolist())
            
            if not suppliers_in_suivi_list:
                st.info("Aucun fournisseur trouvé dans les données de suivi des commandes après traitement.")
            else:
                st.markdown("Sélectionnez les fournisseurs pour lesquels générer un fichier de suivi :")
                selected_fournisseurs_tab5_ui = render_supplier_checkboxes(
                    "tab5", suppliers_in_suivi_list, default_select_all=False
                )

                if not selected_fournisseurs_tab5_ui:
                    st.info("Veuillez sélectionner un ou plusieurs fournisseurs pour générer les fichiers de suivi.")
                else:
                    st.markdown("---")
                    st.markdown(f"**{len(selected_fournisseurs_tab5_ui)} fournisseur(s) sélectionné(s) pour l'export.**")

                    if st.button("📦 Générer et Télécharger les Fichiers de Suivi", key="generate_suivi_btn_main_tab5"):
                        output_cols_suivi_export = ["Date Pièce BC", "N° de pièce", "AF_RefFourniss", "Désignation Article", "Qté Commandées", "Date de livraison prévue"]
                        source_cols_needed_suivi = ["Date Pièce BC", "N° de pièce", "AF_RefFourniss", "Désignation Article", "Qté Commandées", "Fournisseur"]
                        missing_source_cols_suivi_check = [col for col in source_cols_needed_suivi if col not in df_suivi_commandes_all_data.columns]

                        if missing_source_cols_suivi_check:
                            st.error(f"Colonnes sources manquantes ('Suivi commandes'): {', '.join(missing_source_cols_suivi_check)}. Export impossible.")
                        else:
                            export_count_suivi = 0
                            for supplier_name_suivi_export in selected_fournisseurs_tab5_ui:
                                df_supplier_suivi_export_data = df_suivi_commandes_all_data[
                                    df_suivi_commandes_all_data["Fournisseur"] == supplier_name_suivi_export
                                ].copy()

                                if df_supplier_suivi_export_data.empty:
                                    st.warning(f"Aucune commande en cours trouvée pour : {supplier_name_suivi_export}")
                                    continue
                                
                                # Ensure all output columns exist, even if empty initially
                                df_export_final_suivi = pd.DataFrame(columns=output_cols_suivi_export)
                                
                                # Populate with data, handling potential NaT for dates
                                if 'Date Pièce BC' in df_supplier_suivi_export_data:
                                    df_export_final_suivi["Date Pièce BC"] = pd.to_datetime(df_supplier_suivi_export_data["Date Pièce BC"], errors='coerce').dt.strftime('%d/%m/%Y').fillna('')
                                for col_map in ["N° de pièce", "AF_RefFourniss", "Désignation Article", "Qté Commandées"]:
                                     if col_map in df_supplier_suivi_export_data:
                                        df_export_final_suivi[col_map] = df_supplier_suivi_export_data[col_map]
                                
                                df_export_final_suivi["Date de livraison prévue"] = "" # Ensure empty column

                                excel_buffer_suivi_export = io.BytesIO()
                                with pd.ExcelWriter(excel_buffer_suivi_export, engine="openpyxl", date_format='DD/MM/YYYY', datetime_format='DD/MM/YYYY') as writer_suivi_export:
                                    df_export_final_suivi[output_cols_suivi_export].to_excel(writer_suivi_export, sheet_name=sanitize_sheet_name(f"Suivi_{supplier_name_suivi_export}"), index=False)
                                excel_buffer_suivi_export.seek(0)
                                
                                file_name_suivi_export = f"Suivi_Commande_{sanitize_sheet_name(supplier_name_suivi_export)}_{pd.Timestamp.now():%Y%m%d}.xlsx"
                                
                                st.download_button(
                                    label=f"📥 Télécharger Suivi pour {supplier_name_suivi_export}",
                                    data=excel_buffer_suivi_export,
                                    file_name=file_name_suivi_export,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key=f"dl_suivi_{sanitize_supplier_key(supplier_name_suivi_export)}_main_tab5"
                                )
                                export_count_suivi +=1
                            if export_count_suivi > 0: st.success(f"{export_count_suivi} fichier(s) de suivi prêt(s) au téléchargement.")
                            else: st.info("Aucun fichier de suivi généré.")


# --- App Footer / Initial Message if no file is loaded ---
elif not uploaded_file:
    st.info("👋 Bienvenue ! Chargez votre fichier Excel principal pour démarrer l'analyse et les prévisions.")
    if st.button("🔄 Réinitialiser l'Application (efface toutes les données en session)"):
        keys_to_clear_full_reset_app = list(st.session_state.keys())
        for key_to_del_app in keys_to_clear_full_reset_app: del st.session_state[key_to_del_app]
        logging.info("Application state fully reset by user.")
        st.rerun()
elif 'df_initial_filtered' in st.session_state and not isinstance(st.session_state.df_initial_filtered, pd.DataFrame):
    # This case indicates a problem during the loading of df_initial_filtered.
    st.error("Erreur interne : L'état des données filtrées initiales est invalide. Veuillez recharger le fichier.")
    st.session_state.df_full = None # Force re-evaluation of the new file upload block
    if st.button("Réessayer de charger le fichier"):
        st.rerun()
