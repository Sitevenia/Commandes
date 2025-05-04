import streamlit as st
import pandas as pd
import numpy as np
import io
import logging
import re
import openpyxl # Required for engine and direct manipulation
from openpyxl.utils import get_column_letter # Utility to get column letters
import calendar # For month names
# from openpyxl.styles import Font # Uncomment if applying bold font formatting

# --- Logging Configuration ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- Helper Functions ---

def safe_read_excel(uploaded_file, sheet_name, **kwargs):
    """ Safely reads an Excel sheet, returning None if sheet not found or error occurs. """
    try:
        if isinstance(uploaded_file, io.BytesIO): uploaded_file.seek(0)
        file_name = getattr(uploaded_file, 'name', '')
        engine = 'openpyxl' if file_name.lower().endswith('.xlsx') else None
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, engine=engine, **kwargs)
        if df.empty and len(df.columns) == 0:
             logging.warning(f"Sheet '{sheet_name}' was read but appears empty.")
             st.warning(f"⚠️ L'onglet '{sheet_name}' semble vide ou n'a pas d'en-tête valide.")
             return None
        return df
    except ValueError as e:
        if f"Worksheet named '{sheet_name}' not found" in str(e) or f"'{sheet_name}' not found" in str(e):
             logging.warning(f"Sheet '{sheet_name}' not found.")
             st.warning(f"⚠️ Onglet '{sheet_name}' non trouvé.")
        else:
             logging.error(f"ValueError reading sheet '{sheet_name}': {e}")
             st.error(f"❌ Erreur de valeur lecture onglet '{sheet_name}': {e}.")
        return None
    except FileNotFoundError:
        logging.error(f"FileNotFoundError reading sheet '{sheet_name}'.")
        st.error(f"❌ Fichier non trouvé (erreur interne) lecture '{sheet_name}'.")
        return None
    except Exception as e:
        if "zip file" in str(e).lower():
             logging.error(f"Error reading sheet '{sheet_name}': Bad zip file - {e}")
             st.error(f"❌ Erreur lecture onglet '{sheet_name}': Fichier .xlsx corrompu (erreur zip).")
        else:
            logging.error(f"Unexpected error reading sheet '{sheet_name}': {type(e).__name__} - {e}")
            st.error(f"❌ Erreur inattendue ({type(e).__name__}) lecture '{sheet_name}': {e}.")
        return None

def calculer_quantite_a_commander(df, semaine_columns, montant_minimum_input, duree_semaines):
    """ Calcule la quantité à commander. """
    try:
        # Validation
        if not isinstance(df, pd.DataFrame) or df.empty: return None
        required_cols = ["Stock", "Conditionnement", "Tarif d'achat"] + semaine_columns; missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols: st.error(f"Colonnes manquantes calc: {', '.join(missing_cols)}"); return None
        if not semaine_columns: st.error("Colonnes semaines vides calc."); return None
        df_calc = df.copy();
        for col in required_cols: df_calc[col] = pd.to_numeric(df_calc[col], errors='coerce').replace([np.inf, -np.inf], np.nan).fillna(0)
        # Ventes Moyennes
        num_semaines_totales = len(semaine_columns); ventes_N1 = df_calc[semaine_columns].sum(axis=1)
        if num_semaines_totales >= 64: v12N1 = df_calc[semaine_columns[-64:-52]].sum(axis=1); v12N1s = df_calc[semaine_columns[-52:-40]].sum(axis=1); avg12N1 = v12N1 / 12; avg12N1s = v12N1s / 12
        else: v12N1 = pd.Series(0, index=df_calc.index); v12N1s = pd.Series(0, index=df_calc.index); avg12N1 = 0; avg12N1s = 0
        nb_semaines_recentes = min(num_semaines_totales, 12)
        if nb_semaines_recentes > 0: v12last = df_calc[semaine_columns[-nb_semaines_recentes:]].sum(axis=1); avg12last = v12last / nb_semaines_recentes
        else: v12last = pd.Series(0, index=df_calc.index); avg12last = 0
        # Qte Pondérée & Nécessaire
        qpond = (0.5 * avg12last + 0.2 * avg12N1 + 0.3 * avg12N1s); qnec = qpond * duree_semaines
        qcomm_series = (qnec - df_calc["Stock"]).apply(lambda x: max(0, x))
        # Ajustements Règles
        cond = df_calc["Conditionnement"]; stock = df_calc["Stock"]; tarif = df_calc["Tarif d'achat"]; qcomm = qcomm_series.tolist()
        for i in range(len(qcomm)): # Cond
            c = cond.iloc[i]; q = qcomm[i]
            if q > 0 and c > 0: qcomm[i] = int(np.ceil(q / c) * c)
            elif q > 0: qcomm[i] = 0
            else: qcomm[i] = 0
        if nb_semaines_recentes > 0: # R1
            for i in range(len(qcomm)):
                c = cond.iloc[i]; vr_count = (df_calc[semaine_columns[-nb_semaines_recentes:]].iloc[i] > 0).sum()
                if vr_count >= 2 and stock.iloc[i] <= 1 and c > 0: qcomm[i] = max(qcomm[i], c)
        for i in range(len(qcomm)): # R2
            vt_n1 = ventes_N1.iloc[i]; vr_sum = v12last.iloc[i]
            if vt_n1 < 6 and vr_sum < 2: qcomm[i] = 0
        # Ajustement Montant Min
        mt_avant = sum(q * p for q, p in zip(qcomm, tarif))
        if montant_minimum_input > 0 and mt_avant < montant_minimum_input:
            mt_actuel = mt_avant; indices = [i for i, q in enumerate(qcomm) if q > 0]; idx_ptr = 0; max_iter = len(df_calc) * 10; iters = 0
            while mt_actuel < montant_minimum_input and iters < max_iter:
                iters += 1;
                if not indices: break
                curr_idx = indices[idx_ptr % len(indices)]; c = cond.iloc[curr_idx]; p = tarif.iloc[curr_idx]
                if c > 0 and p > 0: qcomm[curr_idx] += c; mt_actuel += c * p
                elif c <= 0 : indices.pop(idx_ptr % len(indices));
                if not indices: continue; idx_ptr -= 1
                idx_ptr += 1
            if iters >= max_iter and mt_actuel < montant_minimum_input: st.error("Ajustement montant min échoué (max iter).")
        # Montant Final
        mt_final = sum(q * p for q, p in zip(qcomm, tarif))
        # Rename calculated columns for clarity in results df
        # Return tuple: quantities, N1_total, N1_sim_12, N_last_12, total_amount
        return (qcomm, ventes_N1, v12N1, v12last, mt_final)
    except Exception as e: st.error(f"Erreur calcul qté: {e}"); logging.exception("Calc Error:"); return None

def calculer_rotation_stock(df, semaine_columns, periode_semaines):
    """ Calcule les métriques de rotation de stock. """
    try:
        if not isinstance(df, pd.DataFrame) or df.empty: return pd.DataFrame()
        required_cols = ["Stock", "Tarif d'achat"];
        if not all(col in df.columns for col in required_cols): missing = [col for col in required_cols if col not in df.columns]; st.error(f"Colonnes manquantes rotation: {', '.join(missing)}"); return None
        df_rotation = df.copy()
        if periode_semaines and periode_semaines > 0 and len(semaine_columns) >= periode_semaines: semaines_analyse = semaine_columns[-periode_semaines:]; nb_semaines_analyse = periode_semaines
        elif periode_semaines and periode_semaines > 0: semaines_analyse = semaine_columns; nb_semaines_analyse = len(semaine_columns); st.caption(f"Analyse sur {nb_semaines_analyse} sem. disponibles.")
        else: semaines_analyse = semaine_columns; nb_semaines_analyse = len(semaine_columns)
        if not semaines_analyse: st.warning("Aucune colonne vente pour analyse rotation."); return df_rotation
        for col in semaines_analyse: df_rotation[col] = pd.to_numeric(df_rotation[col], errors='coerce').fillna(0)
        df_rotation["Unités Vendues (Période)"] = df_rotation[semaines_analyse].sum(axis=1)
        df_rotation["Ventes Moy Hebdo (Période)"] = df_rotation["Unités Vendues (Période)"] / nb_semaines_analyse if nb_semaines_analyse > 0 else 0.0
        avg_weeks_per_month = 52 / 12; df_rotation["Ventes Moy Mensuel (Période)"] = df_rotation["Ventes Moy Hebdo (Période)"] * avg_weeks_per_month
        df_rotation["Stock"] = pd.to_numeric(df_rotation["Stock"], errors='coerce').fillna(0)
        df_rotation["Tarif d'achat"] = pd.to_numeric(df_rotation["Tarif d'achat"], errors='coerce').fillna(0)
        denom_wos = df_rotation["Ventes Moy Hebdo (Période)"]; df_rotation["Semaines Stock (WoS)"] = np.divide(df_rotation["Stock"], denom_wos, out=np.full_like(df_rotation["Stock"], np.inf, dtype=np.float64), where=denom_wos!=0); df_rotation.loc[df_rotation["Stock"] <= 0, "Semaines Stock (WoS)"] = 0.0
        denom_rot_unit = df_rotation["Stock"]; df_rotation["Rotation Unités (Proxy)"] = np.divide(df_rotation["Unités Vendues (Période)"], denom_rot_unit, out=np.full_like(denom_rot_unit, np.inf, dtype=np.float64), where=denom_rot_unit!=0); df_rotation["Rotation Unités (Proxy)"].fillna(0, inplace=True); df_rotation.loc[(df_rotation["Unités Vendues (Période)"] <= 0) & (denom_rot_unit <= 0), "Rotation Unités (Proxy)"] = 0.0
        df_rotation["COGS (Période)"] = df_rotation["Unités Vendues (Période)"] * df_rotation["Tarif d'achat"]; df_rotation["Valeur Stock Actuel (€)"] = df_rotation["Stock"] * df_rotation["Tarif d'achat"]; denom_rot_val = df_rotation["Valeur Stock Actuel (€)"]; df_rotation["Rotation Valeur (Proxy)"] = np.divide(df_rotation["COGS (Période)"], denom_rot_val, out=np.full_like(denom_rot_val, np.inf, dtype=np.float64), where=denom_rot_val!=0); df_rotation["Rotation Valeur (Proxy)"].fillna(0, inplace=True); df_rotation.loc[(df_rotation["COGS (Période)"] <= 0) & (denom_rot_val <= 0), "Rotation Valeur (Proxy)"] = 0.0
        return df_rotation
    except KeyError as e: st.error(f"Erreur clé calc rotation: '{e}'."); return None
    except Exception as e: st.error(f"Erreur inattendue calc rotation: {e}"); logging.exception("Error calc rotation:"); return None

def approx_weeks_to_months(week_columns_52):
    """Approximates month mapping for 52 consecutive week columns."""
    month_map = {}; weeks_per_month_approx = 52 / 12
    for i in range(1, 13):
        start_idx = int(round((i-1) * weeks_per_month_approx)); end_idx = int(round(i * weeks_per_month_approx))
        month_cols = week_columns_52[start_idx : min(end_idx, 52)]; month_name = calendar.month_name[i]; month_map[month_name] = month_cols
    logging.info(f"Approx month map created. Example Jan: {month_map.get('January', [])}")
    return month_map

def calculer_forecast_simulation(df, all_semaine_columns, selected_months, sim_type, progression_pct=0, objectif_montant=0):
    """ Performs forecast simulation for SELECTED MONTHS based on corresponding N-1 data. """
    try:
        if not isinstance(df, pd.DataFrame) or df.empty: st.warning("Aucune donnée pour simulation."); return None
        if len(all_semaine_columns) < 104: st.error("Données historiques insuffisantes (< 104 semaines) pour N-1."); return None
        if not selected_months: st.warning("Veuillez sélectionner au moins un mois."); return None
        required_cols = ["Référence Article", "Désignation Article", "Conditionnement", "Tarif d'achat"]
        if not all(col in df.columns for col in required_cols): missing = [col for col in required_cols if col not in df.columns]; st.error(f"Colonnes manquantes simulation : {', '.join(missing)}"); return None

        df_sim = df[required_cols + ["Fournisseur"]].copy()
        df_sim["Tarif d'achat"] = pd.to_numeric(df_sim["Tarif d'achat"], errors='coerce').fillna(0)
        df_sim["Conditionnement"] = pd.to_numeric(df_sim["Conditionnement"], errors='coerce').fillna(1).apply(lambda x: 1 if x<=0 else int(x))
        n1_week_cols = all_semaine_columns[-104:-52]
        logging.info(f"Forecast Sim: Using N-1 columns from index -104 to -52.")
        if not all(col in df.columns for col in n1_week_cols): st.error("Erreur interne: Colonnes N-1 manquantes."); return None
        df_n1_sales = df[n1_week_cols].copy()
        for col in n1_week_cols: df_n1_sales[col] = pd.to_numeric(df_n1_sales[col], errors='coerce').fillna(0)
        month_col_map_n1 = approx_weeks_to_months(n1_week_cols)
        total_n1_sales_selected_months = pd.Series(0.0, index=df_sim.index); monthly_sales_n1_selected = {}
        for month in selected_months:
            if month in month_col_map_n1 and month_col_map_n1[month]:
                month_n1_cols = [col for col in month_col_map_n1[month] if col in df_n1_sales.columns]
                if month_n1_cols: sales_this_month = df_n1_sales[month_n1_cols].sum(axis=1); monthly_sales_n1_selected[month] = sales_this_month; total_n1_sales_selected_months += sales_this_month; df_sim[f"Ventes N-1 {month}"] = sales_this_month
                else: monthly_sales_n1_selected[month] = pd.Series(0.0, index=df_sim.index); df_sim[f"Ventes N-1 {month}"] = 0.0
            else: logging.warning(f"Aucune colonne N-1 pour mois approx. : {month}"); monthly_sales_n1_selected[month] = pd.Series(0.0, index=df_sim.index); df_sim[f"Ventes N-1 {month}"] = 0.0
        df_sim["Ventes N-1 Total (Mois Sélectionnés)"] = total_n1_sales_selected_months
        period_seasonality = {}; safe_total_n1_sales_selected = total_n1_sales_selected_months.replace(0, np.nan)
        for month in selected_months:
            if month in monthly_sales_n1_selected: period_seasonality[month] = (monthly_sales_n1_selected[month] / safe_total_n1_sales_selected).fillna(0)
            else: period_seasonality[month] = 0.0
        base_monthly_forecast_qty = {}
        if sim_type == 'Simple Progression':
            prog_factor = 1 + (progression_pct / 100.0); total_forecast_qty_selected_period = total_n1_sales_selected_months * prog_factor
            for month in selected_months: base_monthly_forecast_qty[month] = total_forecast_qty_selected_period * period_seasonality.get(month, 0.0)
        elif sim_type == 'Objectif Montant':
            if objectif_montant <= 0: st.error("Objectif montant > 0 requis."); return None
            logging.info(f"Forecast 'Objectif Montant': Distributing Objective={objectif_montant:.2f} based on N-1 Qty Seasonality.")
            total_n1_sales_check = total_n1_sales_selected_months.sum()
            if total_n1_sales_check <= 0:
                st.warning("Ventes N-1 nulles pour mois sélectionnés. Répartition égale tentée.")
                num_selected_months = len(selected_months);
                if num_selected_months == 0: return None
                amount_per_month = objectif_montant / num_selected_months
                for month in selected_months: base_monthly_forecast_qty[month] = np.divide(amount_per_month, df_sim["Tarif d'achat"], out=np.zeros_like(df_sim["Tarif d'achat"], dtype=float), where=df_sim["Tarif d'achat"]!=0)
            else:
                for month in selected_months:
                    target_amount_this_month = objectif_montant * period_seasonality.get(month, 0.0)
                    base_monthly_forecast_qty[month] = np.divide(target_amount_this_month, df_sim["Tarif d'achat"], out=np.zeros_like(df_sim["Tarif d'achat"], dtype=float), where=df_sim["Tarif d'achat"]!=0)
        else: st.error("Type simulation non reconnu."); return None
        total_adjusted_qty = pd.Series(0.0, index=df_sim.index); total_final_amount = pd.Series(0.0, index=df_sim.index)
        for month in selected_months:
            month_qty_col = f"Qté Prév. {month}"; month_amt_col = f"Montant Prév. {month} (€)"
            if month in base_monthly_forecast_qty:
                 base_qty = pd.to_numeric(base_monthly_forecast_qty[month], errors='coerce').fillna(0); cond = df_sim["Conditionnement"]
                 adjusted_qty = (np.ceil(np.divide(base_qty, cond, out=np.zeros_like(base_qty, dtype=float), where=cond!=0)) * cond).fillna(0).astype(int)
                 df_sim[month_qty_col] = adjusted_qty; df_sim[month_amt_col] = adjusted_qty * df_sim["Tarif d'achat"]
                 total_adjusted_qty += adjusted_qty; total_final_amount += df_sim[month_amt_col]
            else: df_sim[month_qty_col] = 0; df_sim[month_amt_col] = 0.0
        df_sim["Qté Totale Prév. (Mois Sélectionnés)"] = total_adjusted_qty; df_sim["Montant Total Prév. (€) (Mois Sélectionnés)"] = total_final_amount
        id_cols = ["Fournisseur", "Référence Article", "Désignation Article", "Conditionnement", "Tarif d'achat"]
        n1_cols = [f"Ventes N-1 {m}" for m in selected_months if f"Ventes N-1 {m}" in df_sim.columns]
        qty_cols = [f"Qté Prév. {m}" for m in selected_months]; amt_cols = [f"Montant Prév. {m} (€)" for m in selected_months]
        total_cols = ["Ventes N-1 Total (Mois Sélectionnés)", "Qté Totale Prév. (Mois Sélectionnés)", "Montant Total Prév. (€) (Mois Sélectionnés)"]
        final_cols = id_cols + total_cols + n1_cols + qty_cols + amt_cols; final_cols = [col for col in final_cols if col in df_sim.columns]
        return df_sim[final_cols]
    except Exception as e: st.error(f"Erreur simulation forecast : {e}"); logging.exception("Error forecast sim calc:"); return None


def sanitize_sheet_name(name):
    """ Removes invalid characters for Excel sheet names and truncates. """
    if not isinstance(name, str): name = str(name)
    sanitized = re.sub(r'[\[\]:*?/\\<>|"]', '_', name)
    if sanitized.startswith("'"): sanitized = "_" + sanitized[1:]
    if sanitized.endswith("'"): sanitized = sanitized[:-1] + "_"
    return sanitized[:31]

# --- Streamlit App Main Logic ---
st.set_page_config(page_title="Forecast & Rotation App", layout="wide")
st.title("📦 Application Prévision Commande & Analyse Rotation")

# --- File Upload ---
uploaded_file = st.file_uploader("📁 Charger le fichier Excel principal", type=["xlsx", "xls"], key="fileUploader")

# Initialize variables / state
if 'df_full' not in st.session_state: st.session_state.df_full = None
if 'min_order_dict' not in st.session_state: st.session_state.min_order_dict = {}
if 'df_initial_filtered' not in st.session_state: st.session_state.df_initial_filtered = pd.DataFrame()
if 'semaine_columns' not in st.session_state: st.session_state.semaine_columns = []
if 'calculation_result_df' not in st.session_state: st.session_state.calculation_result_df = None
if 'rotation_result_df' not in st.session_state: st.session_state.rotation_result_df = None
if 'forecast_result_df' not in st.session_state: st.session_state.forecast_result_df = None
if 'selected_fournisseurs_session' not in st.session_state: st.session_state.selected_fournisseurs_session = []
if 'rotation_threshold_value' not in st.session_state: st.session_state.rotation_threshold_value = 1.0
if 'show_all_rotation' not in st.session_state: st.session_state.show_all_rotation = True
if 'forecast_selected_months' not in st.session_state: st.session_state.forecast_selected_months = list(calendar.month_name)[1:]
if 'forecast_sim_type_index' not in st.session_state: st.session_state.forecast_sim_type_index = 0
if 'forecast_prog_pct' not in st.session_state: st.session_state.forecast_prog_pct = 5.0
if 'forecast_target_amount' not in st.session_state: st.session_state.forecast_target_amount = 10000.0

# --- Data Loading and Initial Processing ---
if uploaded_file and st.session_state.df_full is None:
    logging.info(f"New file uploaded: {uploaded_file.name}. Processing...")
    keys_to_clear_on_new_file = ['df_initial_filtered', 'semaine_columns', 'calculation_result_df', 'rotation_result_df', 'forecast_result_df']
    for key in keys_to_clear_on_new_file:
        if key in st.session_state: del st.session_state[key]
    try:
        file_buffer = io.BytesIO(uploaded_file.getvalue())
        st.info("Lecture onglet 'Tableau final'...")
        df_full_temp = safe_read_excel(file_buffer, sheet_name="Tableau final", header=7)
        if df_full_temp is None: st.error("❌ Échec lecture 'Tableau final'."); st.stop()
        required_on_load = ["Stock", "Fournisseur", "AF_RefFourniss", "Tarif d'achat", "Conditionnement"]; missing_on_load = [col for col in required_on_load if col not in df_full_temp.columns]
        if missing_on_load: st.error(f"❌ Colonnes manquantes: {', '.join(missing_on_load)}"); st.stop()
        df_full_temp["Stock"] = pd.to_numeric(df_full_temp["Stock"], errors='coerce').fillna(0)
        df_full_temp["Tarif d'achat"] = pd.to_numeric(df_full_temp["Tarif d'achat"], errors='coerce').fillna(0)
        df_full_temp["Conditionnement"] = pd.to_numeric(df_full_temp["Conditionnement"], errors='coerce').fillna(1).apply(lambda x: 1 if x<=0 else int(x))
        st.session_state.df_full = df_full_temp; st.success("✅ Onglet 'Tableau final' lu.")
        st.info("Lecture onglet 'Minimum de commande'...")
        df_min_commande_temp = safe_read_excel(file_buffer, sheet_name="Minimum de commande"); min_order_dict_temp = {}
        if df_min_commande_temp is not None:
            st.success("✅ Onglet 'Minimum de commande' lu.")
            supplier_col_min = "Fournisseur"; min_amount_col = "Minimum de Commande; required_min_cols = [supplier_col_min, min_amount_col]
            if all(col in df_min_commande_temp.columns for col in required_min_cols):
                try:
                    df_min_commande_temp[supplier_col_min] = df_min_commande_temp[supplier_col_min].astype(str).str.strip(); df_min_commande_temp[min_amount_col] = pd.to_numeric(df_min_commande_temp[min_amount_col], errors='coerce')
                    min_order_dict_temp = df_min_commande_temp.dropna(subset=[supplier_col_min, min_amount_col]).set_index(supplier_col_min)[min_amount_col].to_dict()
                except Exception as e_min_proc: st.error(f"❌ Erreur traitement 'Min commande': {e_min_proc}")
            else: st.warning(f"⚠️ Colonnes manquantes ({', '.join(required_min_cols)}) dans 'Min commande'.")
        st.session_state.min_order_dict = min_order_dict_temp
        df = st.session_state.df_full
        try:
            filter_cols = ["Fournisseur", "AF_RefFourniss"];
            if not all(col in df.columns for col in filter_cols): st.error(f"❌ Colonnes filtrage ({', '.join(filter_cols)}) manquantes."); st.stop()
            df_init_filtered = df[(df["Fournisseur"].notna()) & (df["Fournisseur"] != "") & (df["Fournisseur"] != "#FILTER") & (df["AF_RefFourniss"].notna()) & (df["AF_RefFourniss"] != "")].copy()
            st.session_state.df_initial_filtered = df_init_filtered
            start_col_index = 12; semaine_cols_temp = []
            if len(df.columns) > start_col_index:
                potential_week_cols = df.columns[start_col_index:].tolist(); exclude_cols = ["Tarif d'achat", "Conditionnement", "Stock", "Total", "Stock à terme", "Ventes N-1", "Ventes 12 semaines identiques N-1", "Ventes 12 dernières semaines", "Quantité à commander", "Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article"]
                semaine_cols_temp = [col for col in potential_week_cols if col not in exclude_cols and pd.api.types.is_numeric_dtype(df.get(col, pd.Series(dtype=float)).dtype)]
            st.session_state.semaine_columns = semaine_cols_temp
            if not semaine_cols_temp: logging.warning("No week columns identified.")
            essential_numeric_cols = ["Stock", "Conditionnement", "Tarif d'achat"]; missing_essential = False
            for col in essential_numeric_cols:
                 if col in df_init_filtered.columns: df_init_filtered[col] = pd.to_numeric(df_init_filtered[col], errors='coerce').fillna(0)
                 elif not df_init_filtered.empty: st.error(f"Colonne essentielle '{col}' manquante."); missing_essential = True
            if missing_essential: st.stop()
            st.rerun()
        except KeyError as e_filter: st.error(f"❌ Colonne filtrage '{e_filter}' manquante."); st.stop()
        except Exception as e_filter_other: st.error(f"❌ Erreur filtrage initial : {e_filter_other}"); st.stop()
    except Exception as e_load: st.error(f"❌ Erreur lecture fichier : {e_load}"); logging.exception("File loading error:"); st.stop()

# --- Main App UI ---
if 'df_initial_filtered' in st.session_state and st.session_state.df_initial_filtered is not None:
    df_full = st.session_state.df_full
    df_base_filtered = st.session_state.get('df_initial_filtered', pd.DataFrame())
    fournisseurs_list = sorted(df_base_filtered["Fournisseur"].unique().tolist()) if not df_base_filtered.empty and "Fournisseur" in df_base_filtered.columns else []
    min_order_dict = st.session_state.min_order_dict
    semaine_columns = st.session_state.semaine_columns

    st.sidebar.header("Filtres (pour Prévision & Rotation)")
    selected_fournisseurs = st.sidebar.multiselect("👤 Fournisseur(s)", options=fournisseurs_list, default=st.session_state.selected_fournisseurs_session, key="supplier_select_sidebar", disabled=not bool(fournisseurs_list))
    st.session_state.selected_fournisseurs_session = selected_fournisseurs
    if selected_fournisseurs:
        df_display_filtered = df_base_filtered[df_base_filtered["Fournisseur"].isin(selected_fournisseurs)].copy()
        if df_display_filtered.empty and fournisseurs_list: st.sidebar.warning("Aucun article trouvé.")
        elif not df_display_filtered.empty: st.sidebar.info(f"{len(df_display_filtered)} articles sélectionnés.")
    else:
        df_display_filtered = df_base_filtered.copy()
        if not selected_fournisseurs and fournisseurs_list: st.sidebar.info("Affichage tous fournisseurs.")

    # --- Tabs ---
    tab1, tab2, tab3, tab4 = st.tabs(["Prévision Commande", "Analyse Rotation Stock", "Vérification Stock", "Simulation Forecast"])

    # ========================= TAB 1: Prévision Commande =========================
    with tab1:
        st.header("Prévision Quantités à Commander"); st.caption("Utilise fournisseurs sélectionnés.")
        # CORRECTED Condition block structure
        if df_display_filtered.empty:
             if selected_fournisseurs:
                 st.warning("Aucun article trouvé pour le(s) fournisseur(s) sélectionné(s).")
             else:
                 if not fournisseurs_list:
                      st.info("Aucun fournisseur valide trouvé dans le fichier après filtrage initial.")
                 else:
                      st.info("Veuillez sélectionner au moins un fournisseur dans la barre latérale.")
        elif not semaine_columns:
            st.warning("Impossible de calculer: Aucune colonne de ventes valide identifiée.")
        else:
            st.markdown("#### Paramètres"); col1_cmd, col2_cmd = st.columns(2)
            with col1_cmd: duree_semaines_cmd = st.number_input(label="⏳ Durée couverture (sem.)", min_value=1, max_value=260, value=4, step=1, key="duree_cmd")
            with col2_cmd: montant_minimum_input_cmd = st.number_input(label="💶 Montant min global (€)", min_value=0.0, max_value=1e12, value=0.0, step=50.0, format="%.2f", key="montant_min_cmd")
            if st.button("🚀 Calculer Quantités", key="calc_cmd_btn"):
                with st.spinner("Calcul..."): result_cmd = calculer_quantite_a_commander(df_display_filtered, semaine_columns, montant_minimum_input_cmd, duree_semaines_cmd)
                if result_cmd:
                    st.success("✅ Calcul OK."); (q_calc, vN1, v12N1, v12l, mt_calc) = result_cmd; df_res_cmd = df_display_filtered.copy()
                    df_res_cmd["Qte Cmdée"] = q_calc; df_res_cmd["Vts N-1 Total"] = vN1; df_res_cmd["Vts 12 N-1 Sim"] = v12N1; df_res_cmd["Vts 12 Dern."] = v12l
                    df_res_cmd["Tarif Ach."] = pd.to_numeric(df_res_cmd["Tarif d'achat"], errors='coerce').fillna(0); df_res_cmd["Total Cmd"] = df_res_cmd["Tarif Ach."] * df_res_cmd["Qte Cmdée"]; df_res_cmd["Stock Terme"] = df_res_cmd["Stock"] + df_res_cmd["Qte Cmdée"]
                    st.session_state.calc_res_df = df_res_cmd; st.session_state.mt_calc = mt_calc; st.session_state.sel_fourn_calc_cmd = selected_fournisseurs; st.rerun()
                else: st.error("❌ Calcul échoué.");
            if 'calc_res_df' in st.session_state and st.session_state.calc_res_df is not None:
                if st.session_state.sel_fourn_calc_cmd == selected_fournisseurs:
                    st.markdown("---"); st.markdown("#### Résultats Commande"); df_cmd_disp = st.session_state.calc_res_df; mt_cmd_disp = st.session_state.mt_calc; sup_cmd_disp = st.session_state.sel_fourn_calc_cmd
                    st.metric(label="💰 Montant Total", value=f"{mt_cmd_disp:,.2f} €")
                    # CORRECTED Indentation Min Warning
                    if len(sup_cmd_disp) == 1:
                        sup_cmd = sup_cmd_disp[0]
                        if sup_cmd in min_order_dict:
                            req_min = min_order_dict[sup_cmd]
                            if "Total Cmd" in df_cmd_disp.columns:
                                act_tot = df_cmd_disp["Total Cmd"].sum()
                                if req_min > 0 and act_tot < req_min:
                                    diff = req_min - act_tot; st.warning(f"⚠️ **Min Non Atteint ({sup_cmd})**\nMontant: **{act_tot:,.2f}€** | Requis: **{req_min:,.2f}€** (Manque: {diff:,.2f}€)")
                            else: logging.warning("Col 'Total Cmd' absente.")
                    # Display Table
                    cols_req = ["Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article", "Stock"]; cols_base = cols_req + ["Vts N-1 Total", "Vts 12 N-1 Sim", "Vts 12 Dern.", "Conditionnement", "Qte Cmdée", "Stock Terme", "Tarif Ach.", "Total Cmd"]
                    cols_disp = [c for c in cols_base if c in df_cmd_disp.columns];
                    if any(c not in df_cmd_disp.columns for c in cols_req): st.error("❌ Cols manquantes affichage.")
                    else: st.dataframe(df_cmd_disp[cols_disp].style.format({"Tarif Ach.": "{:,.2f}€", "Total Cmd": "{:,.2f}€", "Vts N-1 Total": "{:,.0f}", "Vts 12 N-1 Sim": "{:,.0f}", "Vts 12 Dern.": "{:,.0f}", "Stock": "{:,.0f}", "Conditionnement": "{:,.0f}", "Qte Cmdée": "{:,.0f}", "Stock Terme": "{:,.0f}"}, na_rep="-", thousands=","))
                    st.markdown("#### Export Commande"); df_exp_cmd = df_cmd_disp[df_cmd_disp["Qte Cmdée"] > 0].copy()
                    if not df_exp_cmd.empty:
                         out_cmd = io.BytesIO(); sheets_cr_cmd = 0
                         try: # Export logic
                             with pd.ExcelWriter(out_cmd, engine="openpyxl") as writer_cmd:
                                 qty_c, price_c, tot_c = "Qte Cmdée", "Tarif Ach.", "Total Cmd"; export_cols_cmd = [c for c in cols_disp if c != 'Fournisseur']; formula_ok = False
                                 if all(c in export_cols_cmd for c in [qty_c, price_c, tot_c]):
                                     try: qty_idx, price_idx, tot_idx = export_cols_cmd.index(qty_c), export_cols_cmd.index(price_c), export_cols_cmd.index(tot_c); qty_l, price_l, tot_l = get_column_letter(qty_idx + 1), get_column_letter(price_idx + 1), get_column_letter(tot_idx + 1); formula_ok = True
                                     except: pass
                                 if formula_ok:
                                     for sup_exp in sup_cmd_disp:
                                         df_sup_exp = df_exp_cmd[df_exp_cmd["Fournisseur"] == sup_exp].copy();
                                         if not df_sup_exp.empty:
                                             df_sh_data = df_sup_exp[export_cols_cmd].copy(); n_rows = len(df_sh_data); tot_v = df_sh_data[tot_c].sum(); req_m = min_order_dict.get(sup_exp, 0); min_f = f"{req_m:,.2f}€" if req_m > 0 else "N/A"
                                             lbl_c = "Désignation Article" if "Désignation Article" in export_cols_cmd else export_cols_cmd[1]; tot_r = {c: "" for c in export_cols_cmd}; tot_r[lbl_c] = "TOTAL"; tot_r[tot_c] = tot_v; min_r = {c: "" for c in export_cols_cmd}; min_r[lbl_c] = "Min Requis"; min_r[tot_c] = min_f
                                             df_sh = pd.concat([df_sh_data, pd.DataFrame([tot_r]), pd.DataFrame([min_r])], ignore_index=True); s_name = sanitize_sheet_name(sup_exp)
                                             try: # Write sheet and formulas
                                                 df_sh.to_excel(writer_cmd, sheet_name=s_name, index=False); ws = writer_cmd.sheets[s_name];
                                                 # Formulas for data rows
                                                 for r in range(2, n_rows + 2): form = f"={qty_l}{r}*{price_l}{r}"; cell = ws[f"{tot_l}{r}"]; cell.value = form; cell.number_format = '#,##0.00€'
                                                 # SUM formula (Corrected Indentation)
                                                 total_formula_row_cmd = n_rows + 2
                                                 if n_rows > 0: sum_form = f"=SUM({tot_l}2:{tot_l}{n_rows + 1})"; s_cell = ws[f"{tot_l}{total_formula_row_cmd}"]; s_cell.value = sum_form; s_cell.number_format = '#,##0.00€'
                                                 # Increment count (Corrected Indentation)
                                                 sheets_cr_cmd += 1
                                             except Exception as we: logging.exception(f"Err sheet {s_name}:{we}")
                                 else: st.error("Export CMD: Erreur identification colonnes formules.")
                         except Exception as e_w: logging.exception(f"Err writer:{e_w}")
                         if sheets_cr_cmd > 0:
                              out_cmd.seek(0); fname = f"commande_{'multi' if len(sup_cmd_disp)>1 else sanitize_sheet_name(sup_cmd_disp[0])}_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx"
                              st.download_button(f"📥 Télécharger ({sheets_cr_cmd})", out_cmd, fname, key="dl_cmd_btn")
                    else: st.info("Aucune qté > 0 à exporter.")
                else: st.info("Résultats précédents invalidés. Relancez calcul.")

    # ====================== TAB 2: Analyse Rotation Stock ======================
    with tab2:
        # ... (Tab 2 code remains unchanged) ...
        st.header("Analyse Rotation Stocks"); st.caption("Utilise fournisseurs sélectionnés.")
        if not selected_fournisseurs: st.info("Sélectionnez fournisseur(s).")
        elif df_display_filtered.empty and selected_fournisseurs: st.warning("Aucun article trouvé.")
        elif not semaine_columns: st.warning("Colonnes ventes manquantes.")
        else:
            st.markdown("#### Paramètres"); col1_r, col2_r = st.columns(2)
            with col1_r: period_opts = {"12 sem.": 12, "52 sem.": 52, "Total": 0}; sel_p_lbl = st.selectbox("📅 Période:", period_opts.keys(), key="rot_p_sel"); sel_p_w = period_opts[sel_p_lbl]
            with col2_r: st.markdown("##### Options Affichage"); show_all = st.checkbox("Afficher tout", value=st.session_state.show_all_rotation, key="show_all_rot_cb"); st.session_state.show_all_rotation = show_all; rot_thr = st.number_input("... ou ventes mens. <", 0.0, value=st.session_state.rotation_threshold_value, step=0.1, format="%.1f", key="rot_thr_in", disabled=show_all)
            if not show_all: st.session_state.rotation_threshold_value = rot_thr
            if st.button("🔄 Analyser Rotation", key="analyze_rot_btn"):
                 with st.spinner("Analyse..."): df_rot_res = calculer_rotation_stock(df_display_filtered, semaine_columns, sel_p_w)
                 if df_rot_res is not None: st.success("✅ Analyse terminée."); st.session_state.rot_res_df = df_rot_res; st.session_state.rot_p_lbl = sel_p_lbl; st.session_state.sel_fourn_calc_rot = selected_fournisseurs; st.rerun()
                 else: st.error("❌ Analyse échouée.");
            if 'rot_res_df' in st.session_state and st.session_state.rot_res_df is not None:
                 if st.session_state.sel_fourn_calc_rot == selected_fournisseurs:
                    st.markdown("---"); st.markdown(f"#### Résultats Rotation ({st.session_state.get('rot_p_lbl', '')})"); df_rot_orig = st.session_state.rot_res_df; thr_disp = st.session_state.rotation_threshold_value; show_all_f = st.session_state.show_all_rotation
                    m_sales_col = "Ventes Moy Mensuel (Période)"; can_filt = False; df_rot_disp = pd.DataFrame()
                    if m_sales_col in df_rot_orig.columns: m_sales_ser = pd.to_numeric(df_rot_orig[m_sales_col], errors='coerce').fillna(0); can_filt = True
                    else: st.warning(f"Col '{m_sales_col}' non trouvée.")
                    # CORRECTED Filter block structure
                    if show_all_f:
                        df_rot_disp = df_rot_orig.copy(); st.caption(f"Affichage {len(df_rot_disp)} articles.")
                    elif can_filt:
                        try:
                            df_rot_disp = df_rot_orig[m_sales_ser < thr_disp].copy()
                            st.caption(f"Filtre: Ventes < {thr_disp:.1f}/mois. {len(df_rot_disp)} / {len(df_rot_orig)} articles.")
                        except Exception as ef:
                             st.error(f"Err filtre: {ef}")
                             df_rot_disp = df_rot_orig.copy() # Fallback on error
                    else: # Cannot filter
                        df_rot_disp = df_rot_orig.copy()
                        if not show_all_f: # Only warn if filtering was intended
                            st.info("Affichage tous résultats car filtrage impossible.")

                    cols_rot = ["AF_RefFourniss", "Référence Article", "Désignation Article", "Tarif d'achat", "Stock", "Unités Vendues (Période)", "Ventes Moy Hebdo (Période)", "Ventes Moy Mensuel (Période)", "Semaines Stock (WoS)", "Rotation Unités (Proxy)", "Valeur Stock Actuel (€)", "COGS (Période)", "Rotation Valeur (Proxy)"]
                    cols_rot_fin = [c for c in cols_rot if c in df_rot_disp.columns]
                    if df_rot_disp.empty:
                        if not df_rot_orig.empty and not show_all_f and can_filt: st.info(f"Aucun article < {thr_disp:.1f} ventes/mois.")
                    elif not cols_rot_fin: st.error("Cols rotation manquantes.")
                    else:
                        df_rot_disp_cp = df_rot_disp[cols_rot_fin].copy(); num_round = {"Tarif d'achat": 2, "Ventes Moy Hebdo (Période)": 2, "Ventes Moy Mensuel (Période)": 2, "Semaines Stock (WoS)": 1, "Rotation Unités (Proxy)": 2, "Valeur Stock Actuel (€)": 2, "COGS (Période)": 2, "Rotation Valeur (Proxy)": 2}
                        for c, d in num_round.items():
                            if c in df_rot_disp_cp.columns: df_rot_disp_cp[c] = pd.to_numeric(df_rot_disp_cp[c], errors='coerce');
                            if pd.api.types.is_numeric_dtype(df_rot_disp_cp[c]): df_rot_disp_cp[c] = df_rot_disp_cp[c].round(d)
                        df_rot_disp_cp.replace([np.inf, -np.inf], 'Inf', inplace=True); fmters = {"Tarif d'achat": "{:,.2f}€", "Stock": "{:,.0f}", "Unités Vendues (Période)": "{:,.0f}", "Ventes Moy Hebdo (Période)": "{:,.2f}", "Ventes Moy Mensuel (Période)": "{:,.2f}", "Semaines Stock (WoS)": "{}", "Rotation Unités (Proxy)": "{}", "Valeur Stock Actuel (€)": "{:,.2f}€", "COGS (Période)": "{:,.2f}€", "Rotation Valeur (Proxy)": "{}"}
                        st.dataframe(df_rot_disp_cp.style.format(fmters, na_rep="-", thousands=","))
                    st.markdown("#### Export Analyse Affichée")
                    if not df_rot_disp.empty:
                         out_r = io.BytesIO(); cols_r_base = ["AF_RefFourniss", "Référence Article", "Désignation Article", "Tarif d'achat", "Stock", "Unités Vendues (Période)", "Ventes Moy Hebdo (Période)", "Ventes Moy Mensuel (Période)", "Semaines Stock (WoS)", "Rotation Unités (Proxy)", "Valeur Stock Actuel (€)", "COGS (Période)", "Rotation Valeur (Proxy)"]
                         cols_r_fourn = ["Fournisseur"] + cols_r_base if "Fournisseur" in df_rot_disp.columns else cols_r_base; cols_r_fin = [c for c in cols_r_fourn if c in df_rot_disp.columns]
                         df_exp_r = df_rot_disp[cols_r_fin].copy();
                         for c, d in num_round.items():
                              if c in df_exp_r.columns: df_exp_r[c] = pd.to_numeric(df_exp_r[c], errors='coerce');
                              if pd.api.types.is_numeric_dtype(df_exp_r[c]): df_exp_r[c] = df_exp_r[c].round(d)
                         df_exp_r.replace([np.inf, -np.inf], 'Infini', inplace=True); lbl_exp = f"Filtree_{thr_disp:.1f}" if not show_all_f else "Complete"; sh_name = f"Rotation_{lbl_exp}"; f_base = f"analyse_rotation_{lbl_exp}"
                         with pd.ExcelWriter(out_r, engine="openpyxl") as wr_r: df_exp_r.to_excel(wr_r, sheet_name=sh_name, index=False)
                         out_r.seek(0); sups_exp = st.session_state.get('selected_fournisseurs_session', []); f_rot = f"{f_base}_{'multi' if len(sups_exp)>1 else sanitize_sheet_name(sups_exp[0] if sups_exp else 'NA')}_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx"
                         dl_lbl = f"📥 Télécharger {'Filtrée' if not show_all_f else 'Complète'}" + (f" (<{thr_disp:.1f}/m)" if not show_all_f else ""); st.download_button(dl_lbl, out_r, f_rot, key="dl_rot_btn")
                    elif not df_rot_orig.empty: st.info(f"Aucune donnée selon critères (<{thr_disp:.1f}/m) à exporter.")
                    else: st.info("Aucune donnée à exporter.")
                 else: st.info("Résultats analyse invalidés. Relancez.")

    # ========================= TAB 3: Vérification Stock =========================
    with tab3:
        # ... (Tab 3 code remains unchanged) ...
        st.header("Vérification Stocks Négatifs"); st.caption("Analyse tous articles du fichier.")
        df_neg_src = st.session_state.get('df_full', None)
        if df_neg_src is None: st.warning("Données non chargées.")
        elif df_neg_src.empty: st.warning("Aucune donnée dans 'Tableau final'.")
        else:
            stock_c = "Stock";
            if stock_c not in df_neg_src.columns: st.error(f"Colonne '{stock_c}' non trouvée.")
            else:
                df_neg = df_neg_src[df_neg_src[stock_c] < 0].copy()
                if df_neg.empty: st.success("✅ Aucun stock négatif.")
                else:
                    st.warning(f"⚠️ **{len(df_neg)} article(s) avec stock négatif !**"); cols_neg = ["Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article", "Stock"]; cols_neg_fin = [c for c in cols_neg if c in df_neg.columns]
                    if not cols_neg_fin: st.error("Cols manquantes affichage.")
                    else: st.dataframe(df_neg[cols_neg_fin].style.format({"Stock": "{:,.0f}"}, na_rep="-").apply(lambda x: ['background-color:#FADBD8' if v<0 else '' for v in x], subset=['Stock']))
                    st.markdown("---"); st.markdown("#### Exporter Stocks Négatifs"); out_neg = io.BytesIO(); df_exp_n = df_neg[cols_neg_fin].copy()
                    try:
                        with pd.ExcelWriter(out_neg, engine="openpyxl") as w_neg: df_exp_n.to_excel(w_neg, sheet_name="Stocks_Negatifs", index=False)
                        out_neg.seek(0); f_neg = f"stocks_negatifs_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx"; st.download_button("📥 Télécharger Liste", out_neg, f_neg, key="dl_neg_btn")
                    except Exception as e_exp_n: st.error(f"Err export neg: {e_exp_n}");

    # ========================= TAB 4: Simulation Forecast =========================
    with tab4:
        st.header("Simulation Forecast Annuel"); st.caption("Utilise fournisseurs sélectionnés & suppose N-1 = sem. -104 à -52."); st.warning("🚨 **Approximation Importante:** Saisonnalité mensuelle basée sur découpage approx. des 52 sem. N-1.")
        if df_display_filtered.empty:
             if selected_fournisseurs: st.warning("Aucun article trouvé."); else: st.info("Sélectionnez fournisseur(s).")
        elif len(semaine_columns) < 104: st.warning("Données historiques insuffisantes (< 104 sem).")
        else:
            st.markdown("#### Paramètres")
            all_months = list(calendar.month_name)[1:]; default_months = st.session_state.get('forecast_selected_months', all_months); sel_months_fcst = st.multiselect("📅 Mois simulation:", all_months, default_months, key="fcst_months_sel"); st.session_state.forecast_selected_months = sel_months_fcst
            sim_t = st.radio("⚙️ Type Simulation:", ('Simple Progression', 'Objectif Montant'), key="fcst_sim_type", horizontal=True, index=st.session_state.get('forecast_sim_type_index', 0)); st.session_state.forecast_sim_type_index = 0 if sim_t == 'Simple Progression' else 1
            prog_pct = 0.0; obj_mt = 0.0; col1_f, col2_f = st.columns(2)
            with col1_f:
                if sim_t == 'Simple Progression': prog_pct = st.number_input(label="📈 Progression (%)", min_value=-100.0, value=st.session_state.get('forecast_prog_pct', 5.0), step=0.5, format="%.1f", key="fcst_prog_pct")
            with col2_f:
                if sim_t == 'Objectif Montant': obj_mt = st.number_input(label="🎯 Objectif Montant (€) (pour mois sélectionnés)", min_value=0.0, value=st.session_state.get('forecast_target_amount', 10000.0), step=1000.0, format="%.2f", key="fcst_target_amount")

            if st.button("▶️ Lancer Simulation", key="run_fcst_sim"):
                 if not sel_months_fcst: st.warning("Sélectionnez mois.")
                 else:
                    curr_prog = st.session_state.get('forecast_prog_pct', 5.0); curr_obj = st.session_state.get('forecast_target_amount', 10000.0); prog_use = curr_prog if sim_t == 'Simple Progression' else 0; obj_use = curr_obj if sim_t == 'Objectif Montant' else 0
                    with st.spinner("Simulation..."): df_fcst_res = calculer_forecast_simulation(df_display_filtered, semaine_columns, sel_months_fcst, sim_t, prog_use, obj_use)
                    if df_fcst_res is not None: st.success("✅ Simulation terminée."); st.session_state.forecast_result_df = df_fcst_res; st.session_state.forecast_params = {'suppliers': selected_fournisseurs, 'months': sel_months_fcst, 'type': sim_t, 'prog': prog_use, 'obj': obj_use}; st.rerun()
                    else: st.error("❌ Simulation échouée.");
            if 'forecast_result_df' in st.session_state and st.session_state.forecast_result_df is not None:
                 current_params_disp = {'suppliers': selected_fournisseurs, 'months': sel_months_fcst, 'type': sim_t, 'prog': st.session_state.get('forecast_prog_pct', 5.0) if sim_t=='Simple Progression' else 0, 'obj': st.session_state.get('forecast_target_amount', 10000.0) if sim_t=='Objectif Montant' else 0}
                 if st.session_state.get('forecast_params') == current_params_disp:
                    st.markdown("---"); st.markdown("#### Résultats Simulation")
                    df_fcst_disp = st.session_state.forecast_result_df;
                    mq_cols = [f"Qté Prév. {m}" for m in sel_months_fcst if f"Qté Prév. {m}" in df_fcst_disp.columns]; ma_cols = [f"Montant Prév. {m} (€)" for m in sel_months_fcst if f"Montant Prév. {m} (€)" in df_fcst_disp.columns]; n1m_cols = [f"Ventes N-1 {m}" for m in sel_months_fcst if f"Ventes N-1 {m}" in df_fcst_disp.columns]
                    fcst_id = ["Fournisseur", "Référence Article", "Désignation Article", "Conditionnement", "Tarif d'achat"]; fcst_tot = ["Ventes N-1 Total (Mois Sélectionnés)", "Qté Totale Prév. (Mois Sélectionnés)", "Montant Total Prév. (€) (Mois Sélectionnés)"]
                    fcst_disp_cols = fcst_id + fcst_tot + n1m_cols + mq_cols + ma_cols; fcst_disp_fin = [c for c in fcst_disp_cols if c in df_fcst_disp.columns]
                    if df_fcst_disp.empty: st.info("Aucun résultat.")
                    elif not fcst_disp_fin: st.error("Erreur: Colonnes résultats manquantes.")
                    else:
                        fcst_fmters = {"Tarif d'achat": "{:,.2f}€", "Conditionnement": "{:,.0f}", "Ventes N-1 Total (Mois Sélectionnés)": "{:,.0f}", "Qté Totale Prév. (Mois Sélectionnés)": "{:,.0f}", "Montant Total Prév. (€) (Mois Sélectionnés)": "{:,.2f}€"}
                        for c in n1m_cols: fcst_fmters[c] = "{:,.0f}";
                        for c in mq_cols: fcst_fmters[c] = "{:,.0f}";
                        for c in ma_cols: fcst_fmters[c] = "{:,.2f}€"
                        fcst_fmters_final = {k: v for k, v in fcst_fmters.items() if k in fcst_disp_fin}
                        try: st.dataframe(df_fcst_disp[fcst_disp_fin].style.format(fcst_fmters_final, na_rep="-", thousands=","))
                        except Exception as e_fmt: st.error(f"Erreur formatage affichage: {e_fmt}"); st.dataframe(df_fcst_disp[fcst_disp_fin])
                        st.markdown("#### Export Simulation"); out_f = io.BytesIO(); df_exp_f = df_fcst_disp[fcst_disp_fin].copy()
                        try:
                            with pd.ExcelWriter(out_f, engine="openpyxl") as w_f: df_exp_f.to_excel(w_f, sheet_name=f"Forecast_{sim_t.replace(' ','_')}", index=False)
                            out_f.seek(0); fb = f"forecast_{sim_t.replace(' ','_').lower()}"; sups_f = st.session_state.get('selected_fournisseurs_session', [])
                            f_fcst = f"{fb}_{'multi' if len(sups_f)>1 else sanitize_sheet_name(sups_f[0] if sups_f else 'NA')}_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx"; st.download_button("📥 Télécharger Simulation", out_f, f_fcst, key="dl_fcst_btn")
                        except Exception as eef: st.error(f"Err export forecast: {eef}")
                 else: st.info("Résultats simulation invalidés. Relancez.")


# --- App footer/initial message ---
elif not uploaded_file:
    st.info("👋 Bienvenue ! Chargez votre fichier Excel pour commencer.")
    if st.button("🔄 Réinitialiser l'application"):
         keys_to_clear = list(st.session_state.keys())
         for key in keys_to_clear: del st.session_state[key]
         st.rerun() # Use st.rerun()
