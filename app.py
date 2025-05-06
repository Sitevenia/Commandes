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
        
        if df is None: # Should not happen if read_excel doesn't raise error, but defensive
            logging.error(f"Pandas read_excel returned None for sheet '{sheet_name}' without raising an error.")
            return None
        logging.debug(f"Read sheet '{sheet_name}'. DataFrame empty: {df.empty}, Columns: {df.columns.tolist()}, Shape: {df.shape}")
        
        if len(df.columns) == 0:
             logging.warning(f"Sheet '{sheet_name}' was read but has no columns.")
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
    except FileNotFoundError:
        logging.error(f"FileNotFoundError (unexpected with BytesIO) reading sheet '{sheet_name}'.")
        st.error(f"❌ Fichier non trouvé (erreur interne) lors de la lecture de l'onglet '{sheet_name}'.")
        return None
    except Exception as e:
        if "zip file" in str(e).lower():
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
        for col in required_cols:
            df_calc[col] = pd.to_numeric(df_calc[col], errors='coerce').replace([np.inf, -np.inf], np.nan).fillna(0)

        num_semaines_totales = len(semaine_columns)
        ventes_N1 = df_calc[semaine_columns].sum(axis=1)

        if num_semaines_totales >= 64:
            v12N1 = df_calc[semaine_columns[-64:-52]].sum(axis=1)
            v12N1s = df_calc[semaine_columns[-52:-40]].sum(axis=1)
            avg12N1 = v12N1 / 12
            avg12N1s = v12N1s / 12
        else:
            v12N1 = pd.Series(0.0, index=df_calc.index)
            v12N1s = pd.Series(0.0, index=df_calc.index)
            avg12N1 = pd.Series(0.0, index=df_calc.index)
            avg12N1s = pd.Series(0.0, index=df_calc.index)

        nb_semaines_recentes = min(num_semaines_totales, 12)
        if nb_semaines_recentes > 0:
            v12last = df_calc[semaine_columns[-nb_semaines_recentes:]].sum(axis=1)
            avg12last = v12last / nb_semaines_recentes
        else:
            v12last = pd.Series(0.0, index=df_calc.index)
            avg12last = pd.Series(0.0, index=df_calc.index)

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
                logging.warning(f"Article index {df_calc.index[i]} (Ref: {df_calc.get('Référence Article', pd.Series(['N/A']))[i]}) Qté {q:.2f} ignorée car conditionnement est {c}.")
                qcomm[i] = 0 
            else:
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
        elif periode_semaines and periode_semaines > 0:
            semaines_analyse = semaine_columns
            nb_semaines_analyse = len(semaine_columns)
            st.caption(f"Période d'analyse ajustée à {nb_semaines_analyse} semaines (données disponibles).")
        else:
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
    month_map = {}
    if not week_columns_52 or len(week_columns_52) != 52:
        logging.warning(f"approx_weeks_to_months expects 52 columns, got {len(week_columns_52) if week_columns_52 else 0}.")
        return month_map

    weeks_per_month_approx = 52 / 12.0
    
    for i in range(1, 13):
        month_name = calendar.month_name[i]
        start_idx = int(round((i-1) * weeks_per_month_approx))
        end_idx = int(round(i * weeks_per_month_approx))
        month_cols = week_columns_52[start_idx : min(end_idx, 52)]
        month_map[month_name] = month_cols

    logging.info(f"Approximated month-to-week map. Jan: {month_map.get('January', [])}")
    return month_map

def calculer_forecast_simulation_v3(df, all_semaine_columns, selected_months, sim_type, progression_pct=0, objectif_montant=0):
    try:
        if not isinstance(df, pd.DataFrame) or df.empty:
            st.warning("Aucune donnée pour simulation forecast.")
            return None, 0.0

        if not all_semaine_columns or len(all_semaine_columns) < 52:
            st.error("Données historiques < 52 semaines pour N-1.")
            return None, 0.0

        if not selected_months:
            st.warning("Veuillez sélectionner au moins un mois pour la simulation.")
            return None, 0.0

        required_cols = ["Référence Article", "Désignation Article", "Conditionnement", "Tarif d'achat", "Fournisseur"]
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            st.error(f"Colonnes manquantes (simulation): {', '.join(missing_cols)}")
            return None, 0.0

        years_in_cols = set()
        parsed_week_cols = []
        for col_name in all_semaine_columns:
            if isinstance(col_name, str):
                match = re.match(r"(\d{4})S?(\d{1,2})", col_name, re.IGNORECASE)
                if match:
                    year, week = int(match.group(1)), int(match.group(2))
                    if 1 <= week <= 53:
                        years_in_cols.add(year)
                        parsed_week_cols.append({'year': year, 'week': week, 'col': col_name, 'sort_key': year * 100 + week})
        
        if not years_in_cols:
            st.error("Impossible de déterminer les années. Format attendu: 'YYYYWW' ou 'YYYYSwW'.")
            return None, 0.0

        parsed_week_cols.sort(key=lambda x: x['sort_key'])
        
        year_n = max(years_in_cols) if years_in_cols else 0
        year_n_minus_1 = year_n - 1
        
        st.caption(f"Simulation N-1 (Année N: {year_n}, Année N-1: {year_n_minus_1})")

        n1_week_cols_data = [item for item in parsed_week_cols if item['year'] == year_n_minus_1]
        
        if len(n1_week_cols_data) < 52:
            st.error(f"Données N-1 ({year_n_minus_1}) insuffisantes: {len(n1_week_cols_data)} sem. trouvées (52 req.).")
            return None, 0.0
        
        n1_week_cols_for_mapping = [item['col'] for item in n1_week_cols_data[:52]]

        df_sim = df[required_cols].copy()
        df_sim["Tarif d'achat"] = pd.to_numeric(df_sim["Tarif d'achat"], errors='coerce').fillna(0)
        df_sim["Conditionnement"] = pd.to_numeric(df_sim["Conditionnement"], errors='coerce').fillna(1).apply(lambda x: 1 if x <= 0 else int(x))

        missing_n1_in_df = [col for col in n1_week_cols_for_mapping if col not in df.columns]
        if missing_n1_in_df:
            st.error(f"Erreur interne: Colonnes N-1 mappées ({', '.join(missing_n1_in_df)}) non trouvées dans DataFrame.")
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
            df_sim[f"Ventes N-1 {month_name}"] = sales_this_month_n1
        
        df_sim["Vts N-1 Tot (Mois Sel.)"] = total_n1_sales_selected_months_series

        period_seasonality_factors = {}
        safe_total_n1_for_selected_months = total_n1_sales_selected_months_series.copy()

        for month_name in selected_months:
            month_sales_n1 = monthly_sales_n1_for_selected_months.get(month_name, pd.Series(0.0, index=df_sim.index))
            factor = np.divide(month_sales_n1, safe_total_n1_for_selected_months, 
                               out=np.zeros_like(month_sales_n1, dtype=float),
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
                st.error("Objectif Montant > 0 requis.")
                return None, 0.0

            total_n1_sales_units_all_items = total_n1_sales_selected_months_series.sum()

            if total_n1_sales_units_all_items <= 0:
                st.warning("Ventes N-1 nulles. Répartition égale du montant objectif / mois / articles (avec tarif > 0).")
                num_sel_months = len(selected_months)
                if num_sel_months == 0: return None, 0.0
                
                target_amt_per_month = objectif_montant / num_sel_months
                num_items_with_price = (df_sim["Tarif d'achat"] > 0).sum()

                for month_name in selected_months:
                    if num_items_with_price == 0:
                        base_monthly_forecast_qty_map[month_name] = pd.Series(0.0, index=df_sim.index)
                    else:
                        target_amt_per_item_this_month = target_amt_per_month / num_items_with_price
                        base_monthly_forecast_qty_map[month_name] = np.divide(
                            target_amt_per_item_this_month, df_sim["Tarif d'achat"],
                            out=np.zeros_like(df_sim["Tarif d'achat"], dtype=float),
                            where=df_sim["Tarif d'achat"] != 0
                        )
            else:
                for month_name in selected_months:
                    seasonality_for_month = period_seasonality_factors.get(month_name, pd.Series(0.0, index=df_sim.index))
                    target_amount_for_this_month_per_item = objectif_montant * seasonality_for_month
                    
                    base_monthly_forecast_qty_map[month_name] = np.divide(
                        target_amount_for_this_month_per_item, df_sim["Tarif d'achat"],
                        out=np.zeros_like(df_sim["Tarif d'achat"], dtype=float),
                        where=df_sim["Tarif d'achat"] != 0
                    )
        else:
            st.error(f"Type de simulation non reconnu: '{sim_type}'.")
            return None, 0.0

        total_adjusted_qty_all_months = pd.Series(0.0, index=df_sim.index)
        total_final_amount_all_months = pd.Series(0.0, index=df_sim.index)

        for month_name in selected_months:
            forecast_qty_col_name = f"Qté Prév. {month_name}"
            forecast_amount_col_name = f"Montant Prév. {month_name} (€)"
            
            base_q_series = base_monthly_forecast_qty_map.get(month_name, pd.Series(0.0, index=df_sim.index))
            base_q_series = pd.to_numeric(base_q_series, errors='coerce').fillna(0)
            cond_series = df_sim["Conditionnement"]
            
            adjusted_qty_series = (
                np.ceil(
                    np.divide(base_q_series, cond_series, 
                              out=np.zeros_like(base_q_series, dtype=float), 
                              where=cond_series != 0)
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
        
        df_sim.rename(columns={
            "Qté Totale Prév. (Mois Sel.)": "Qté Tot Prév (Mois Sel.)",
            "Montant Total Prév. (€) (Mois Sel.)": "Mnt Tot Prév (€) (Mois Sel.)"
        }, inplace=True)
        total_cols_display = [
            "Vts N-1 Tot (Mois Sel.)",
            "Qté Tot Prév (Mois Sel.)",
            "Mnt Tot Prév (€) (Mois Sel.)"
        ]

        final_ordered_cols = id_cols_display + total_cols_display + n1_sales_cols_display + qty_forecast_cols_display + amt_forecast_cols_display
        final_ordered_cols_existing = [col for col in final_ordered_cols if col in df_sim.columns]

        grand_total_forecast_amount = total_final_amount_all_months.sum()
        
        return df_sim[final_ordered_cols_existing], grand_total_forecast_amount

    except KeyError as e:
        st.error(f"Erreur de clé (simulation forecast): '{e}'.")
        logging.exception(f"KeyError in calculer_forecast_simulation_v3: {e}")
        return None, 0.0
    except Exception as e:
        st.error(f"Erreur inattendue (simulation forecast): {type(e).__name__} - {e}")
        logging.exception("Error in calculer_forecast_simulation_v3:")
        return None, 0.0

def sanitize_sheet_name(name):
    if not isinstance(name, str): name = str(name)
    sanitized = re.sub(r'[\[\]:*?/\\<>|"]', '_', name)
    if sanitized.startswith("'"): sanitized = "_" + sanitized[1:]
    if sanitized.endswith("'"): sanitized = sanitized[:-1] + "_"
    return sanitized[:31]

def render_supplier_checkboxes(tab_key_prefix, all_suppliers, default_select_all=False):
    select_all_key = f"{tab_key_prefix}_select_all"
    supplier_cb_keys = {
        supplier: f"{tab_key_prefix}_cb_{sanitize_supplier_key(supplier)}" for supplier in all_suppliers
    }

    if select_all_key not in st.session_state:
        st.session_state[select_all_key] = default_select_all
        for cb_key in supplier_cb_keys.values():
            if cb_key not in st.session_state:
                st.session_state[cb_key] = default_select_all
    else:
        for cb_key in supplier_cb_keys.values():
            if cb_key not in st.session_state:
                 st.session_state[cb_key] = st.session_state[select_all_key]

    def toggle_all_suppliers_for_tab():
        current_select_all_value = st.session_state[select_all_key]
        for cb_key in supplier_cb_keys.values():
            st.session_state[cb_key] = current_select_all_value

    def check_individual_supplier_for_tab():
        all_individual_checked = all(
            st.session_state.get(cb_key, False) for cb_key in supplier_cb_keys.values()
        )
        if st.session_state.get(select_all_key) != all_individual_checked:
            st.session_state[select_all_key] = all_individual_checked
    
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
            checkbox_cols[current_col_idx].checkbox(
                supplier_name,
                key=cb_key,
                on_change=check_individual_supplier_for_tab
            )
            if st.session_state.get(cb_key):
                selected_suppliers_in_ui.append(supplier_name)
            current_col_idx = (current_col_idx + 1) % num_display_cols
    return selected_suppliers_in_ui

def sanitize_supplier_key(supplier_name):
     if not isinstance(supplier_name, str): supplier_name = str(supplier_name)
     s = re.sub(r'\W+', '_', supplier_name)
     s = re.sub(r'^_+|_+$', '', s)
     s = re.sub(r'_+', '_', s)
     return s if s else "invalid_supplier_key"

# --- Streamlit App Main Logic ---
st.set_page_config(page_title="Forecast & Rotation App", layout="wide")
st.title("📦 Application Prévision Commande, Analyse Rotation & Suivi")

uploaded_file = st.file_uploader(
    "📁 Charger le fichier Excel principal (contenant 'Tableau final', 'Minimum de commande', 'Suivi commandes')",
    type=["xlsx", "xls"],
    key="main_file_uploader"
)

def get_default_session_state():
    return {
        'df_full': None,
        'min_order_dict': {},
        'df_initial_filtered': pd.DataFrame(),
        'all_available_semaine_columns': [],
        'unique_suppliers_list': [],
        'commande_result_df': None,
        'commande_calculated_total_amount': 0.0,
        'commande_suppliers_calculated_for': [],
        'rotation_result_df': None,
        'rotation_analysis_period_label': "",
        'rotation_suppliers_calculated_for': [],
        'rotation_threshold_value': 1.0,
        'show_all_rotation_data': True,
        'forecast_result_df': None,
        'forecast_grand_total_amount': 0.0,
        'forecast_simulation_params_calculated_for': {},
        'forecast_selected_months_ui': list(calendar.month_name)[1:],
        'forecast_sim_type_radio_index': 0,
        'forecast_progression_percentage_ui': 5.0,
        'forecast_target_amount_ui': 10000.0,
        'df_suivi_commandes': None,
    }

for key, default_value in get_default_session_state().items():
    if key not in st.session_state:
        st.session_state[key] = default_value

if uploaded_file and st.session_state.df_full is None:
    logging.info(f"New file uploaded: {uploaded_file.name}. Starting processing...")
    
    keys_to_reset_on_new_file = list(get_default_session_state().keys())
    dynamic_key_prefixes_to_clear = ['tab1_', 'tab2_', 'tab3_', 'tab4_', 'tab5_']

    for key in keys_to_reset_on_new_file:
        if key in st.session_state: del st.session_state[key]
    for prefix in dynamic_key_prefixes_to_clear:
        for k_to_remove in [k for k in st.session_state if k.startswith(prefix)]:
            del st.session_state[k_to_remove]
    for key, default_value in get_default_session_state().items():
        st.session_state[key] = default_value
    logging.info("Session state has been reset and re-initialized for the new file.")

    try:
        excel_file_buffer = io.BytesIO(uploaded_file.getvalue())
        
        st.info("Lecture de l'onglet 'Tableau final'...")
        df_full_temp = safe_read_excel(excel_file_buffer, sheet_name="Tableau final", header=7)
        if df_full_temp is None: st.error("❌ Échec lecture 'Tableau final'."); st.stop()

        required_cols_tf = ["Stock", "Fournisseur", "AF_RefFourniss", "Tarif d'achat", "Conditionnement", "Référence Article", "Désignation Article"]
        missing_cols_tf_check = [col for col in required_cols_tf if col not in df_full_temp.columns]
        if missing_cols_tf_check: st.error(f"❌ Cols manquantes ('Tableau final'): {', '.join(missing_cols_tf_check)}."); st.stop()

        df_full_temp["Stock"] = pd.to_numeric(df_full_temp["Stock"], errors='coerce').fillna(0)
        df_full_temp["Tarif d'achat"] = pd.to_numeric(df_full_temp["Tarif d'achat"], errors='coerce').fillna(0)
        df_full_temp["Conditionnement"] = pd.to_numeric(df_full_temp["Conditionnement"], errors='coerce').fillna(1).apply(lambda x: int(x) if x > 0 else 1)
        for str_col in ["Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article"]:
            if str_col in df_full_temp.columns: df_full_temp[str_col] = df_full_temp[str_col].astype(str).str.strip()
        st.session_state.df_full = df_full_temp
        st.success("✅ 'Tableau final' lu.")

        st.info("Lecture 'Minimum de commande'...")
        excel_file_buffer.seek(0)
        df_min_commande_temp = safe_read_excel(excel_file_buffer, sheet_name="Minimum de commande")
        min_order_dict_temp = {}
        if df_min_commande_temp is not None:
            s_col, m_col = "Fournisseur", "Minimum de Commande" # MODIFIED: Shortened var names
            if s_col in df_min_commande_temp.columns and m_col in df_min_commande_temp.columns:
                try:
                    df_min_commande_temp[s_col] = df_min_commande_temp[s_col].astype(str).str.strip()
                    df_min_commande_temp[m_col] = pd.to_numeric(df_min_commande_temp[m_col], errors='coerce')
                    min_order_dict_temp = df_min_commande_temp.dropna(subset=[s_col, m_col]).set_index(s_col)[m_col].to_dict()
                    st.success(f"✅ 'Minimum de commande' lu ({len(min_order_dict_temp)} entrées).")
                except Exception as e_min: st.error(f"❌ Erreur traitement 'Minimum de commande': {e_min}")
            else: st.warning(f"⚠️ Cols '{s_col}'/'{m_col}' manquantes ('Minimum de commande').")
        st.session_state.min_order_dict = min_order_dict_temp

        st.info("Lecture onglet 'Suivi commandes'...")
        excel_file_buffer.seek(0)
        df_suivi_temp = safe_read_excel(excel_file_buffer, sheet_name="Suivi commandes", header=4)
        
        if df_suivi_temp is not None:
            # MODIFIED: required_suivi_cols with "Intitulé Fournisseur"
            required_suivi_cols = ["Date Pièce BC", "N° de pièce", "AF_RefFourniss", "Désignation Article", "Qté Commandées", "Intitulé Fournisseur"]
            missing_suivi_cols_check = [col for col in required_suivi_cols if col not in df_suivi_temp.columns]
            
            if not missing_suivi_cols_check:
                # Rename "Intitulé Fournisseur" to "Fournisseur" for consistency within the app
                df_suivi_temp.rename(columns={"Intitulé Fournisseur": "Fournisseur"}, inplace=True)

                for col_to_strip in ["Fournisseur", "AF_RefFourniss", "Désignation Article", "N° de pièce"]:
                    if col_to_strip in df_suivi_temp.columns: df_suivi_temp[col_to_strip] = df_suivi_temp[col_to_strip].astype(str).str.strip()
                if "Qté Commandées" in df_suivi_temp.columns:
                    df_suivi_temp["Qté Commandées"] = pd.to_numeric(df_suivi_temp["Qté Commandées"], errors='coerce').fillna(0)
                if "Date Pièce BC" in df_suivi_temp.columns:
                    try: df_suivi_temp["Date Pièce BC"] = pd.to_datetime(df_suivi_temp["Date Pièce BC"], errors='coerce')
                    except Exception as e_dt: st.warning(f"⚠️ Problème parsing 'Date Pièce BC' (Suivi): {e_dt}.")
                
                df_suivi_temp.dropna(how='all', inplace=True)
                st.session_state.df_suivi_commandes = df_suivi_temp
                st.success(f"✅ 'Suivi commandes' lu ({len(df_suivi_temp)} lignes).")
            else:
                st.warning(f"⚠️ Cols manquantes ('Suivi commandes', L5): {', '.join(missing_suivi_cols_check)}. Suivi limité.")
                st.session_state.df_suivi_commandes = pd.DataFrame()
        else:
            st.info("Onglet 'Suivi commandes' non trouvé/vide. Suivi non dispo.")
            st.session_state.df_suivi_commandes = pd.DataFrame()

        df_loaded_ff = st.session_state.df_full
        df_init_filtered_temp = df_loaded_ff[
            (df_loaded_ff["Fournisseur"].notna()) & (df_loaded_ff["Fournisseur"] != "") & (df_loaded_ff["Fournisseur"] != "#FILTER") &
            (df_loaded_ff["AF_RefFourniss"].notna()) & (df_loaded_ff["AF_RefFourniss"] != "")
        ].copy()
        st.session_state.df_initial_filtered = df_init_filtered_temp

        f_w_c_idx = 12 # MODIFIED: Shortened var name
        pot_sales_cols = [] # MODIFIED: Shortened var name
        if len(df_loaded_ff.columns) > f_w_c_idx:
            cand_cols_s = df_loaded_ff.columns[f_w_c_idx:].tolist() # MODIFIED: Shortened var name
            known_non_w_cols = ["Tarif d'achat", "Conditionnement", "Stock", "Total", "Stock à terme", "Ventes N-1", "Ventes 12 semaines identiques N-1", "Ventes 12 dernières semaines", "Quantité à commander", "Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article"]
            excl_set = set(known_non_w_cols) # MODIFIED: Shortened var name
            for col_c in cand_cols_s: # MODIFIED: Shortened var name
                if col_c not in excl_set and pd.api.types.is_numeric_dtype(df_loaded_ff.get(col_c, pd.Series(dtype=object)).dtype):
                    pot_sales_cols.append(col_c)
        st.session_state.all_available_semaine_columns = pot_sales_cols
        if not pot_sales_cols: st.warning("⚠️ Aucune colonne de vente numérique identifiée.")

        if not df_init_filtered_temp.empty:
            st.session_state.unique_suppliers_list = sorted(df_init_filtered_temp["Fournisseur"].astype(str).unique().tolist())
        
        st.rerun()

    except Exception as e_load_main:
        st.error(f"❌ Erreur majeure chargement/traitement initial: {e_load_main}")
        logging.exception("Major file loading/processing error:")
        st.session_state.df_full = None 
        st.stop()

if 'df_initial_filtered' in st.session_state and isinstance(st.session_state.df_initial_filtered, pd.DataFrame):
    df_base_for_tabs = st.session_state.df_initial_filtered
    all_suppliers_from_data = st.session_state.unique_suppliers_list
    min_order_amounts = st.session_state.min_order_dict
    identified_semaine_cols = st.session_state.all_available_semaine_columns
    df_suivi_commandes_all_data = st.session_state.get('df_suivi_commandes', pd.DataFrame())

    tab_titles = ["Prévision Commande", "Analyse Rotation Stock", "Vérification Stock", "Simulation Forecast", "Suivi Commandes Fourn."]
    tab1, tab2, tab3, tab4, tab5 = st.tabs(tab_titles)

    with tab1: # Prévision Commande
        st.header("Prévision des Quantités à Commander")
        selected_fournisseurs_tab1 = render_supplier_checkboxes("tab1", all_suppliers_from_data, default_select_all=True)
        df_display_tab1 = pd.DataFrame() 
        if selected_fournisseurs_tab1:
            if not df_base_for_tabs.empty:
                df_display_tab1 = df_base_for_tabs[df_base_for_tabs["Fournisseur"].isin(selected_fournisseurs_tab1)].copy()
                st.caption(f"{len(df_display_tab1)} art. / {len(selected_fournisseurs_tab1)} fourn.")
        else: st.info("Sélectionner au moins un fournisseur.")
        st.markdown("---")
        if df_display_tab1.empty and selected_fournisseurs_tab1 :
            st.warning("Aucun article pour fournisseur(s) sélectionné(s).")
        elif not identified_semaine_cols and not df_display_tab1.empty:
            st.warning("Colonnes ventes (semaines) non identifiées.")
        elif not df_display_tab1.empty :
            st.markdown("#### Paramètres Calcul Commande")
            col1_cmd, col2_cmd = st.columns(2)
            with col1_cmd: duree_sem_cmd = st.number_input("⏳ Couverture (sem.)", 1, 260, 4, 1, key="duree_cmd_ui_t1")
            with col2_cmd: mt_min_cmd = st.number_input("💶 Montant min global (€)", 0.0, value=0.0, step=50.0, format="%.2f", key="mt_min_cmd_ui_t1") # MODIFIED: Key
            
            if st.button("🚀 Calculer Quantités Cmd", key="calc_qte_cmd_btn_t1"): # MODIFIED: Key
                with st.spinner("Calcul quantités..."):
                    res_cmd = calculer_quantite_a_commander(df_display_tab1, identified_semaine_cols, mt_min_cmd, duree_sem_cmd)
                if res_cmd:
                    st.success("✅ Calcul quantités OK.")
                    q_calc, vN1, v12N1, v12l, mt_calc = res_cmd
                    df_res_cmd = df_display_tab1.copy()
                    df_res_cmd["Qte Cmdée"] = q_calc
                    df_res_cmd["Vts N-1 Total (calc)"] = vN1; df_res_cmd["Vts 12 N-1 Sim (calc)"] = v12N1; df_res_cmd["Vts 12 Dern. (calc)"] = v12l
                    df_res_cmd["Tarif Ach."] = pd.to_numeric(df_res_cmd["Tarif d'achat"], errors='coerce').fillna(0)
                    df_res_cmd["Total Cmd (€)"] = df_res_cmd["Tarif Ach."] * df_res_cmd["Qte Cmdée"]
                    df_res_cmd["Stock Terme"] = df_res_cmd["Stock"] + df_res_cmd["Qte Cmdée"]
                    st.session_state.commande_result_df = df_res_cmd
                    st.session_state.commande_calculated_total_amount = mt_calc
                    st.session_state.commande_suppliers_calculated_for = selected_fournisseurs_tab1
                    st.rerun()
                else: st.error("❌ Calcul quantités échoué.")

            if st.session_state.commande_result_df is not None and st.session_state.commande_suppliers_calculated_for == selected_fournisseurs_tab1:
                st.markdown("---"); st.markdown("#### Résultats Prévision Commande")
                df_cmd_disp = st.session_state.commande_result_df; mt_cmd_disp = st.session_state.commande_calculated_total_amount; sup_cmd_disp = st.session_state.commande_suppliers_calculated_for
                st.metric(label="💰 Montant Total Commandé", value=f"{mt_cmd_disp:,.2f} €")
                if len(sup_cmd_disp) == 1:
                    sup_s = sup_cmd_disp[0]
                    if sup_s in min_order_amounts:
                        req_min_s = min_order_amounts[sup_s]; act_tot_s = df_cmd_disp[df_cmd_disp["Fournisseur"] == sup_s]["Total Cmd (€)"].sum()
                        if req_min_s > 0 and act_tot_s < req_min_s: st.warning(f"⚠️ Min non atteint ({sup_s}): {act_tot_s:,.2f}€ / Requis: {req_min_s:,.2f}€ (Manque: {req_min_s - act_tot_s:,.2f}€)")
                cols_show_cmd = ["Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article", "Stock", "Vts N-1 Total (calc)", "Vts 12 N-1 Sim (calc)", "Vts 12 Dern. (calc)", "Conditionnement", "Qte Cmdée", "Stock Terme", "Tarif Ach.", "Total Cmd (€)"]
                disp_cols_cmd = [c for c in cols_show_cmd if c in df_cmd_disp.columns]
                if not disp_cols_cmd: st.error("Aucune colonne à afficher (commande).")
                else:
                    fmts_cmd = {"Tarif Ach.": "{:,.2f}€", "Total Cmd (€)": "{:,.2f}€", "Vts N-1 Total (calc)": "{:,.0f}", "Vts 12 N-1 Sim (calc)": "{:,.0f}", "Vts 12 Dern. (calc)": "{:,.0f}", "Stock": "{:,.0f}", "Conditionnement": "{:,.0f}", "Qte Cmdée": "{:,.0f}", "Stock Terme": "{:,.0f}"}
                    st.dataframe(df_cmd_disp[disp_cols_cmd].style.format(fmts_cmd, na_rep="-", thousands=","))
                st.markdown("#### Export Commandes")
                df_exp_cmd = df_cmd_disp[df_cmd_disp["Qte Cmdée"] > 0].copy()
                if not df_exp_cmd.empty:
                    out_b_cmd = io.BytesIO(); shts_cmd = 0
                    try:
                        with pd.ExcelWriter(out_b_cmd, engine="openpyxl") as writer_cmd:
                            exp_cols_sht_cmd = [c for c in disp_cols_cmd if c != 'Fournisseur']
                            q_c, p_c, t_c = "Qte Cmdée", "Tarif Ach.", "Total Cmd (€)" # MODIFIED: Shortened
                            f_ok = False # MODIFIED: Shortened
                            if all(c in exp_cols_sht_cmd for c in [q_c, p_c, t_c]):
                                try: q_l, p_l, t_l = get_column_letter(exp_cols_sht_cmd.index(q_c)+1), get_column_letter(exp_cols_sht_cmd.index(p_c)+1), get_column_letter(exp_cols_sht_cmd.index(t_c)+1); f_ok=True # MODIFIED
                                except ValueError: pass
                            for sup_exp in sup_cmd_disp:
                                df_s_exp = df_exp_cmd[df_exp_cmd["Fournisseur"] == sup_exp] # MODIFIED
                                if not df_s_exp.empty:
                                    df_w_sht = df_s_exp[exp_cols_sht_cmd].copy() # MODIFIED
                                    n_r = len(df_w_sht) # MODIFIED
                                    lbl_c_sum = "Désignation Article" if "Désignation Article" in exp_cols_sht_cmd else (exp_cols_sht_cmd[1] if len(exp_cols_sht_cmd) > 1 else exp_cols_sht_cmd[0]) # MODIFIED
                                    tot_v_s = df_w_sht[t_c].sum(); min_r_s = min_order_amounts.get(sup_exp,0); min_d_s = f"{min_r_s:,.2f}€" if min_r_s > 0 else "N/A" # MODIFIED
                                    sum_r = pd.DataFrame([{lbl_c_sum: "TOTAL", t_c: tot_v_s}, {lbl_c_sum: "Min Requis Fourn.", t_c: min_d_s}], columns=exp_cols_sht_cmd).fillna('') # MODIFIED
                                    df_fin_s = pd.concat([df_w_sht, sum_r], ignore_index=True) # MODIFIED
                                    s_nm = sanitize_sheet_name(sup_exp) # MODIFIED
                                    try:
                                        df_fin_s.to_excel(writer_cmd, sheet_name=s_nm, index=False)
                                        ws = writer_cmd.sheets[s_nm]
                                        if f_ok and n_r > 0:
                                            for r_idx in range(2, n_r + 2): ws[f"{t_l}{r_idx}"].value = f"={q_l}{r_idx}*{p_l}{r_idx}"; ws[f"{t_l}{r_idx}"].number_format = '#,##0.00€'
                                            ws[f"{t_l}{n_r + 2}"].value = f"=SUM({t_l}2:{t_l}{n_r + 1})"; ws[f"{t_l}{n_r + 2}"].number_format = '#,##0.00€'
                                        shts_cmd += 1
                                    except Exception as e_sht: logging.error(f"Err export sheet {s_nm}: {e_sht}")
                        if shts_cmd > 0:
                            writer_cmd.save(); out_b_cmd.seek(0)
                            fn_cmd = f"commandes_{'multi' if len(sup_cmd_disp)>1 else sanitize_sheet_name(sup_cmd_disp[0])}_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx" # MODIFIED
                            st.download_button(f"📥 Télécharger ({shts_cmd} feuilles)", out_b_cmd, fn_cmd, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_cmd_btn_t1_dl") # MODIFIED: Key
                        else: st.info("Aucune qté > 0 à exporter.")
                    except Exception as e_wrt_cmd: logging.exception(f"Err ExcelWriter cmd: {e_wrt_cmd}"); st.error("Erreur export commandes.")
                else: st.info("Aucun article qté > 0 à exporter.")
            else: st.info("Résultats commande invalidés. Relancer.")

    with tab2: # Analyse Rotation Stock
        st.header("Analyse de la Rotation des Stocks")
        selected_fournisseurs_tab2 = render_supplier_checkboxes("tab2", all_suppliers_from_data, default_select_all=True)
        df_display_tab2 = pd.DataFrame()
        if selected_fournisseurs_tab2:
            if not df_base_for_tabs.empty:
                df_display_tab2 = df_base_for_tabs[df_base_for_tabs["Fournisseur"].isin(selected_fournisseurs_tab2)].copy()
                st.caption(f"{len(df_display_tab2)} art. / {len(selected_fournisseurs_tab2)} fourn.")
        else: st.info("Sélectionner au moins un fournisseur.")
        st.markdown("---")
        if df_display_tab2.empty and selected_fournisseurs_tab2: st.warning("Aucun article pour fournisseur(s) sélectionné(s).")
        elif not identified_semaine_cols and not df_display_tab2.empty: st.warning("Colonnes ventes (semaines) non identifiées.")
        elif not df_display_tab2.empty:
            st.markdown("#### Paramètres Analyse Rotation")
            col1_rot, col2_rot = st.columns(2)
            with col1_rot:
                period_opts_rot = {"12 dern. sem.": 12, "52 dern. sem.": 52, "Total dispo.": 0}
                sel_p_lbl_rot = st.selectbox("⏳ Période analyse:", period_opts_rot.keys(), key="rot_p_sel_ui_t2") # MODIFIED: Key
                sel_p_w_rot = period_opts_rot[sel_p_lbl_rot]
            with col2_rot:
                st.markdown("##### Options Affichage")
                show_all_rot = st.checkbox("Afficher tout", value=st.session_state.show_all_rotation_data, key="show_all_rot_ui_cb_t2") # MODIFIED: Key
                st.session_state.show_all_rotation_data = show_all_rot
                rot_thr_ui = st.number_input("... ou vts mens. <", 0.0, value=st.session_state.rotation_threshold_value, step=0.1, format="%.1f", key="rot_thr_ui_numin_t2", disabled=show_all_rot) # MODIFIED: Key
                if not show_all_rot: st.session_state.rotation_threshold_value = rot_thr_ui

            if st.button("🔄 Analyser Rotation", key="analyze_rot_btn_t2"): # MODIFIED: Key
                with st.spinner("Analyse rotation..."):
                    df_rot_res = calculer_rotation_stock(df_display_tab2, identified_semaine_cols, sel_p_w_rot)
                if df_rot_res is not None:
                    st.success("✅ Analyse rotation OK."); st.session_state.rotation_result_df = df_rot_res
                    st.session_state.rotation_analysis_period_label = sel_p_lbl_rot; st.session_state.rotation_suppliers_calculated_for = selected_fournisseurs_tab2
                    st.rerun()
                else: st.error("❌ Analyse rotation échouée.")
            
            if st.session_state.rotation_result_df is not None and st.session_state.rotation_suppliers_calculated_for == selected_fournisseurs_tab2:
                st.markdown("---"); st.markdown(f"#### Résultats Rotation ({st.session_state.rotation_analysis_period_label})")
                df_rot_orig = st.session_state.rotation_result_df; thr_disp_rot = st.session_state.rotation_threshold_value; show_all_f_rot = st.session_state.show_all_rotation_data
                m_sales_col_rot = "Ventes Moy Mensuel (Période)"; df_rot_disp = pd.DataFrame()
                if df_rot_orig.empty: st.info("Aucune donnée de rotation à afficher.")
                elif show_all_f_rot: df_rot_disp = df_rot_orig.copy(); st.caption(f"Affichage {len(df_rot_disp)} articles.")
                elif m_sales_col_rot in df_rot_orig.columns:
                    try:
                        sales_filter = pd.to_numeric(df_rot_orig[m_sales_col_rot], errors='coerce').fillna(0)
                        df_rot_disp = df_rot_orig[sales_filter < thr_disp_rot].copy()
                        st.caption(f"Filtre: Vts < {thr_disp_rot:.1f}/mois. {len(df_rot_disp)} / {len(df_rot_orig)} art.")
                        if df_rot_disp.empty: st.info(f"Aucun article < {thr_disp_rot:.1f} vts/mois.")
                    except Exception as ef_rot: st.error(f"Err filtre: {ef_rot}"); df_rot_disp = df_rot_orig.copy()
                else: st.warning(f"Col '{m_sales_col_rot}' non trouvée. Affichage tout."); df_rot_disp = df_rot_orig.copy()
                if not df_rot_disp.empty:
                    cols_rot_show = ["Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article", "Tarif d'achat", "Stock", "Unités Vendues (Période)", "Ventes Moy Hebdo (Période)", "Ventes Moy Mensuel (Période)", "Semaines Stock (WoS)", "Rotation Unités (Proxy)", "Valeur Stock Actuel (€)", "COGS (Période)", "Rotation Valeur (Proxy)"]
                    disp_cols_rot = [c for c in cols_rot_show if c in df_rot_disp.columns]
                    df_disp_cp_rot = df_rot_disp[disp_cols_rot].copy()
                    num_round_rot = {"Tarif d'achat": 2, "Ventes Moy Hebdo (Période)": 2, "Ventes Moy Mensuel (Période)": 2, "Semaines Stock (WoS)": 1, "Rotation Unités (Proxy)": 2, "Valeur Stock Actuel (€)": 2, "COGS (Période)": 2, "Rotation Valeur (Proxy)": 2}
                    for c, d in num_round_rot.items():
                        if c in df_disp_cp_rot.columns: df_disp_cp_rot[c] = pd.to_numeric(df_disp_cp_rot[c], errors='coerce').round(d)
                    df_disp_cp_rot.replace([np.inf, -np.inf], 'Infini', inplace=True)
                    fmts_rot = {"Tarif d'achat": "{:,.2f}€", "Stock": "{:,.0f}", "Unités Vendues (Période)": "{:,.0f}", "Ventes Moy Hebdo (Période)": "{:,.2f}", "Ventes Moy Mensuel (Période)": "{:,.2f}", "Semaines Stock (WoS)": "{}", "Rotation Unités (Proxy)": "{}", "Valeur Stock Actuel (€)": "{:,.2f}€", "COGS (Période)": "{:,.2f}€", "Rotation Valeur (Proxy)": "{}"}
                    st.dataframe(df_disp_cp_rot.style.format(fmts_rot, na_rep="-", thousands=","))
                    st.markdown("#### Export Analyse Affichée")
                    out_b_rot = io.BytesIO(); df_exp_rot = df_disp_cp_rot
                    lbl_exp_rot = f"Filtree_{thr_disp_rot:.1f}" if not show_all_f_rot else "Complete"; sh_name_rot = sanitize_sheet_name(f"Rotation_{lbl_exp_rot}"); f_base_rot = f"analyse_rotation_{lbl_exp_rot}"
                    sup_exp_name_rot = 'multi' if len(selected_fournisseurs_tab2)>1 else (sanitize_sheet_name(selected_fournisseurs_tab2[0]) if selected_fournisseurs_tab2 else 'NA')
                    with pd.ExcelWriter(out_b_rot, engine="openpyxl") as wr_rot: df_exp_rot.to_excel(wr_rot, sheet_name=sh_name_rot, index=False)
                    out_b_rot.seek(0); f_rot_exp = f"{f_base_rot}_{sup_exp_name_rot}_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx"
                    dl_lbl_rot = f"📥 Télécharger ({'Filtrée' if not show_all_f_rot else 'Complète'})"
                    st.download_button(dl_lbl_rot, out_b_rot, f_rot_exp, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_rot_btn_t2_dl") # MODIFIED: Key
            else: st.info("Résultats analyse invalidés. Relancer.")

    with tab3: # Vérification Stock
        st.header("Vérification des Stocks Négatifs")
        st.caption("Analyse tous articles du 'Tableau final'.")
        df_full_neg = st.session_state.get('df_full', None)
        if df_full_neg is None or not isinstance(df_full_neg, pd.DataFrame): st.warning("Données non chargées.")
        elif df_full_neg.empty: st.info("'Tableau final' vide.")
        else:
            stock_c_neg = "Stock"
            if stock_c_neg not in df_full_neg.columns: st.error(f"Colonne '{stock_c_neg}' non trouvée.")
            else:
                df_neg_res = df_full_neg[df_full_neg[stock_c_neg] < 0].copy()
                if df_neg_res.empty: st.success("✅ Aucun stock négatif.")
                else:
                    st.warning(f"⚠️ **{len(df_neg_res)} article(s) avec stock négatif !**")
                    cols_neg_show = ["Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article", "Stock"]
                    disp_cols_neg = [c for c in cols_neg_show if c in df_neg_res.columns]
                    if not disp_cols_neg: st.error("Cols manquantes affichage négatifs.")
                    else: st.dataframe(df_neg_res[disp_cols_neg].style.format({"Stock": "{:,.0f}"}, na_rep="-").apply(lambda s: ['background-color:#FADBD8' if s.name == stock_c_neg and val < 0 else '' for val in s], axis=0))
                    st.markdown("---"); st.markdown("#### Exporter Stocks Négatifs")
                    out_b_neg = io.BytesIO()
                    try:
                        with pd.ExcelWriter(out_b_neg, engine="openpyxl") as w_neg: df_neg_res[disp_cols_neg].to_excel(w_neg, sheet_name="Stocks_Negatifs", index=False)
                        out_b_neg.seek(0); f_neg_exp = f"stocks_negatifs_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx"
                        st.download_button("📥 Télécharger Liste Négatifs", out_b_neg, f_neg_exp, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_neg_btn_t3_dl") # MODIFIED: Key
                    except Exception as e_exp_neg: st.error(f"Err export neg: {e_exp_neg}")

    with tab4: # Simulation Forecast
        st.header("Simulation de Forecast Annuel")
        selected_fournisseurs_tab4 = render_supplier_checkboxes("tab4", all_suppliers_from_data, default_select_all=True)
        df_display_tab4 = pd.DataFrame()
        if selected_fournisseurs_tab4:
            if not df_base_for_tabs.empty:
                df_display_tab4 = df_base_for_tabs[df_base_for_tabs["Fournisseur"].isin(selected_fournisseurs_tab4)].copy()
                st.caption(f"{len(df_display_tab4)} art. / {len(selected_fournisseurs_tab4)} fourn.")
        else: st.info("Sélectionner au moins un fournisseur.")
        st.markdown("---"); st.warning("🚨 **Hypothèse:** Saisonnalité mensuelle approx. sur 52 sem. N-1.")
        if df_display_tab4.empty and selected_fournisseurs_tab4: st.warning("Aucun article pour fournisseur(s) sélectionné(s).")
        elif len(identified_semaine_cols) < 52 and not df_display_tab4.empty : st.warning(f"Données historiques < 52 sem ({len(identified_semaine_cols)}). Simulation N-1 impossible.")
        elif not df_display_tab4.empty:
            st.markdown("#### Paramètres Simulation Forecast")
            all_cal_months = list(calendar.month_name)[1:]
            sel_m_fcst_ui = st.multiselect("📅 Mois simulation:", all_cal_months, default=st.session_state.forecast_selected_months_ui, key="fcst_months_sel_ui_t4") # MODIFIED: Key
            st.session_state.forecast_selected_months_ui = sel_m_fcst_ui
            sim_t_opts_fcst = ('Simple Progression', 'Objectif Montant')
            sim_t_fcst_ui = st.radio("⚙️ Type Simulation:", sim_t_opts_fcst, horizontal=True, index=st.session_state.forecast_sim_type_radio_index, key="fcst_sim_type_ui_t4") # MODIFIED: Key
            st.session_state.forecast_sim_type_radio_index = sim_t_opts_fcst.index(sim_t_fcst_ui)
            prog_pct_fcst, obj_mt_fcst = 0.0, 0.0
            col1_f, col2_f = st.columns(2)
            with col1_f:
                if sim_t_fcst_ui == 'Simple Progression':
                    prog_pct_fcst = st.number_input("📈 Progression (%)", -100.0, value=st.session_state.forecast_progression_percentage_ui, step=0.5, format="%.1f", key="fcst_prog_pct_ui_t4") # MODIFIED: Key
                    st.session_state.forecast_progression_percentage_ui = prog_pct_fcst
            with col2_f:
                if sim_t_fcst_ui == 'Objectif Montant':
                    obj_mt_fcst = st.number_input("🎯 Objectif (€) (mois sel.)", 0.0, value=st.session_state.forecast_target_amount_ui, step=1000.0, format="%.2f", key="fcst_target_amt_ui_t4") # MODIFIED: Key
                    st.session_state.forecast_target_amount_ui = obj_mt_fcst
            if st.button("▶️ Lancer Simulation Forecast", key="run_fcst_sim_btn_t4"): # MODIFIED: Key
                if not sel_m_fcst_ui: st.error("Sélectionner au moins un mois.")
                else:
                    with st.spinner("Simulation forecast..."):
                        df_fcst_res, gt_fcst = calculer_forecast_simulation_v3(df_display_tab4, identified_semaine_cols, sel_m_fcst_ui, sim_t_fcst_ui, prog_pct_fcst, obj_mt_fcst)
                    if df_fcst_res is not None:
                        st.success("✅ Simulation forecast OK."); st.session_state.forecast_result_df = df_fcst_res; st.session_state.forecast_grand_total_amount = gt_fcst
                        st.session_state.forecast_simulation_params_calculated_for = {'suppliers': selected_fournisseurs_tab4, 'months': sel_m_fcst_ui, 'type': sim_t_fcst_ui, 'prog_pct': prog_pct_fcst, 'obj_amt': obj_mt_fcst}
                        st.rerun()
                    else: st.error("❌ Simulation forecast échouée.")
            if st.session_state.forecast_result_df is not None:
                curr_p_fcst_ui = {'suppliers': selected_fournisseurs_tab4, 'months': sel_m_fcst_ui, 'type': sim_t_fcst_ui, 'prog_pct': st.session_state.forecast_progression_percentage_ui if sim_t_fcst_ui=='Simple Progression' else 0.0, 'obj_amt': st.session_state.forecast_target_amount_ui if sim_t_fcst_ui=='Objectif Montant' else 0.0} # MODIFIED
                if st.session_state.forecast_simulation_params_calculated_for == curr_p_fcst_ui: # MODIFIED
                    st.markdown("---"); st.markdown("#### Résultats Simulation Forecast")
                    df_fcst_disp = st.session_state.forecast_result_df; gt_fcst_disp = st.session_state.forecast_grand_total_amount
                    if df_fcst_disp.empty: st.info("Aucun résultat simulation.")
                    else:
                        fmts_fcst = {"Tarif d'achat": "{:,.2f}€", "Conditionnement": "{:,.0f}"}
                        for m_disp in sel_m_fcst_ui: # MODIFIED
                            if f"Ventes N-1 {m_disp}" in df_fcst_disp.columns: fmts_fcst[f"Ventes N-1 {m_disp}"] = "{:,.0f}"
                            if f"Qté Prév. {m_disp}" in df_fcst_disp.columns: fmts_fcst[f"Qté Prév. {m_disp}"] = "{:,.0f}"
                            if f"Montant Prév. {m_disp} (€)" in df_fcst_disp.columns: fmts_fcst[f"Montant Prév. {m_disp} (€)"] = "{:,.2f}€"
                        for col_name in ["Vts N-1 Tot (Mois Sel.)", "Qté Tot Prév (Mois Sel.)", "Mnt Tot Prév (€) (Mois Sel.)"]: # Added loop for totals
                            if col_name in df_fcst_disp.columns:
                                fmts_fcst[col_name] = "{:,.0f}" if "Qté" in col_name or "Vts" in col_name else "{:,.2f}€"
                        try: st.dataframe(df_fcst_disp.style.format(fmts_fcst, na_rep="-", thousands=","))
                        except Exception as e_fmt_fcst: st.error(f"Err format affichage: {e_fmt_fcst}"); st.dataframe(df_fcst_disp)
                        st.metric(label="💰 Mnt Total Prévisionnel (€) (mois sel.)", value=f"{gt_fcst_disp:,.2f} €")
                        st.markdown("#### Export Simulation")
                        out_b_fcst = io.BytesIO(); df_exp_fcst = df_fcst_disp.copy()
                        try:
                            sim_t_fn = sim_t_fcst_ui.replace(' ', '_').lower() # MODIFIED
                            with pd.ExcelWriter(out_b_fcst, engine="openpyxl") as w_fcst: df_exp_fcst.to_excel(w_fcst, sheet_name=sanitize_sheet_name(f"Forecast_{sim_t_fn}"), index=False)
                            out_b_fcst.seek(0)
                            sup_exp_name_fcst = 'multi' if len(selected_fournisseurs_tab4)>1 else (sanitize_sheet_name(selected_fournisseurs_tab4[0]) if selected_fournisseurs_tab4 else 'NA')
                            f_fcst_exp = f"forecast_{sim_t_fn}_{sup_exp_name_fcst}_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx"
                            st.download_button("📥 Télécharger Simulation", out_b_fcst, f_fcst_exp, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_fcst_btn_t4_dl") # MODIFIED: Key
                        except Exception as eef_fcst: st.error(f"Err export forecast: {eef_fcst}")
                else: st.info("Résultats simulation invalidés. Relancer.")

    with tab5: # Suivi Commandes Fournisseurs
        st.header("📄 Suivi des Commandes Fournisseurs")

        if df_suivi_commandes_all_data is None or df_suivi_commandes_all_data.empty:
            st.warning("Aucune donnée de suivi de commandes n'a été chargée (onglet 'Suivi commandes' vide/manquant ou erreur de lecture).")
        else:
            suppliers_in_suivi_list = []
            # Ensure 'Fournisseur' column exists after potential rename
            if "Fournisseur" in df_suivi_commandes_all_data.columns:
                suppliers_in_suivi_list = sorted(df_suivi_commandes_all_data["Fournisseur"].astype(str).unique().tolist())
            
            if not suppliers_in_suivi_list:
                st.info("Aucun fournisseur trouvé dans les données de suivi des commandes traitées.")
            else:
                st.markdown("Sélectionnez les fournisseurs pour lesquels générer un fichier de suivi :")
                selected_fournisseurs_tab5_ui = render_supplier_checkboxes("tab5", suppliers_in_suivi_list, default_select_all=False)

                if not selected_fournisseurs_tab5_ui:
                    st.info("Veuillez sélectionner un ou plusieurs fournisseurs pour générer les fichiers de suivi.")
                else:
                    st.markdown("---"); st.markdown(f"**{len(selected_fournisseurs_tab5_ui)} fournisseur(s) sélectionné(s) pour l'export.**")
                    if st.button("📦 Générer et Télécharger les Fichiers de Suivi", key="generate_suivi_btn_t5"): # MODIFIED: Key
                        output_cols_suivi_export = ["Date Pièce BC", "N° de pièce", "AF_RefFourniss", "Désignation Article", "Qté Commandées", "Date de livraison prévue"]
                        # Source columns needed (after potential rename of "Intitulé Fournisseur" to "Fournisseur")
                        source_cols_needed_suivi = ["Date Pièce BC", "N° de pièce", "AF_RefFourniss", "Désignation Article", "Qté Commandées", "Fournisseur"]
                        missing_source_cols_suivi_check = [col for col in source_cols_needed_suivi if col not in df_suivi_commandes_all_data.columns]

                        if missing_source_cols_suivi_check:
                            st.error(f"Colonnes sources manquantes ('Suivi commandes'): {', '.join(missing_source_cols_suivi_check)}. Export impossible.")
                        else:
                            export_count_suivi = 0
                            for sup_name_s_exp in selected_fournisseurs_tab5_ui: # MODIFIED
                                df_sup_s_exp_data = df_suivi_commandes_all_data[df_suivi_commandes_all_data["Fournisseur"] == sup_name_s_exp].copy() # MODIFIED
                                if df_sup_s_exp_data.empty: st.warning(f"Aucune commande pour : {sup_name_s_exp}"); continue
                                
                                df_exp_final_s = pd.DataFrame(columns=output_cols_suivi_export) # MODIFIED
                                if 'Date Pièce BC' in df_sup_s_exp_data: df_exp_final_s["Date Pièce BC"] = pd.to_datetime(df_sup_s_exp_data["Date Pièce BC"], errors='coerce').dt.strftime('%d/%m/%Y').fillna('')
                                for col_map in ["N° de pièce", "AF_RefFourniss", "Désignation Article", "Qté Commandées"]:
                                     if col_map in df_sup_s_exp_data: df_exp_final_s[col_map] = df_sup_s_exp_data[col_map]
                                df_exp_final_s["Date de livraison prévue"] = ""

                                excel_buf_s_exp = io.BytesIO() # MODIFIED
                                with pd.ExcelWriter(excel_buf_s_exp, engine="openpyxl", date_format='DD/MM/YYYY', datetime_format='DD/MM/YYYY') as writer_s_exp: # MODIFIED
                                    df_exp_final_s[output_cols_suivi_export].to_excel(writer_s_exp, sheet_name=sanitize_sheet_name(f"Suivi_{sup_name_s_exp}"), index=False)
                                excel_buf_s_exp.seek(0)
                                file_name_s_exp = f"Suivi_Commande_{sanitize_sheet_name(sup_name_s_exp)}_{pd.Timestamp.now():%Y%m%d}.xlsx" # MODIFIED
                                st.download_button(label=f"📥 Télécharger Suivi pour {sup_name_s_exp}", data=excel_buf_s_exp, file_name=file_name_s_exp, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"dl_suivi_{sanitize_supplier_key(sup_name_s_exp)}_t5_dl") # MODIFIED: Key
                                export_count_suivi +=1
                            if export_count_suivi > 0: st.success(f"{export_count_suivi} fichier(s) de suivi prêt(s).")
                            else: st.info("Aucun fichier de suivi généré.")

elif not uploaded_file:
    st.info("👋 Bienvenue ! Chargez votre fichier Excel principal pour démarrer.")
    if st.button("🔄 Réinitialiser l'Application"):
        for key_to_del in list(st.session_state.keys()): del st.session_state[key_to_del]
        st.rerun()
elif 'df_initial_filtered' in st.session_state and not isinstance(st.session_state.df_initial_filtered, pd.DataFrame):
    st.error("Erreur interne : Données filtrées invalides. Rechargez le fichier.")
    st.session_state.df_full = None
    if st.button("Réessayer"): st.rerun()
