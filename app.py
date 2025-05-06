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
    """ Calcule la quantité à commander pour le Tab 1. """
    try:
        if not isinstance(df, pd.DataFrame) or df.empty: return None
        required_cols = ["Stock", "Conditionnement", "Tarif d'achat"] + semaine_columns; missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols: st.error(f"Colonnes manquantes calc: {', '.join(missing_cols)}"); return None
        if not semaine_columns: st.error("Colonnes semaines vides calc."); return None
        df_calc = df.copy();
        for col in required_cols: df_calc[col] = pd.to_numeric(df_calc[col], errors='coerce').replace([np.inf, -np.inf], np.nan).fillna(0)
        num_semaines_totales = len(semaine_columns); ventes_N1_total = df_calc[semaine_columns].sum(axis=1) # Renommé pour clarté
        if num_semaines_totales >= 64:
            v12N1 = df_calc[semaine_columns[-64:-52]].sum(axis=1); v12N1s = df_calc[semaine_columns[-52:-40]].sum(axis=1); avg12N1 = v12N1 / 12; avg12N1s = v12N1s / 12
            logging.info("Calcul N-1 basé sur les semaines -64 à -40.")
        else: v12N1 = pd.Series(0, index=df_calc.index); v12N1s = pd.Series(0, index=df_calc.index); avg12N1 = 0; avg12N1s = 0; logging.warning(f"Moins de 64 semaines disponibles ({num_semaines_totales}), calculs N-1 mis à 0.")
        nb_semaines_recentes = min(num_semaines_totales, 12)
        if nb_semaines_recentes > 0:
             v12last = df_calc[semaine_columns[-nb_semaines_recentes:]].sum(axis=1); avg12last = v12last / nb_semaines_recentes # Diviser par le nombre réel
             logging.info(f"Calcul ventes récentes basé sur les {nb_semaines_recentes} dernières semaines.")
        else: v12last = pd.Series(0, index=df_calc.index); avg12last = 0; logging.warning("Aucune semaine récente disponible.")
        qpond = (0.5 * avg12last + 0.2 * avg12N1 + 0.3 * avg12N1s); qnec = qpond * duree_semaines
        qcomm_series = (qnec - df_calc["Stock"]).apply(lambda x: max(0, x))
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
            vt_n1 = ventes_N1_total.iloc[i]; vr_sum = v12last.iloc[i] # v12last est la somme des nb_semaines_recentes
            if vt_n1 < 6 and vr_sum < 2: qcomm[i] = 0
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
        mt_final = sum(q * p for q, p in zip(qcomm, tarif))
        return (qcomm, ventes_N1_total, v12N1, v12last, mt_final) # Renvoyer les ventes calculées
    except Exception as e: st.error(f"Erreur calcul qté: {e}"); logging.exception("Calc Error:"); return None


def calculer_rotation_stock(df, semaine_columns, periode_semaines):
    """ Calcule les métriques de rotation de stock. """
    # --- (Code inchangé) ---
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
    # --- (Code inchangé) ---
    month_map = {}; weeks_per_month_approx = 52 / 12
    for i in range(1, 13):
        start_idx = int(round((i-1) * weeks_per_month_approx)); end_idx = int(round(i * weeks_per_month_approx))
        month_cols = week_columns_52[start_idx : min(end_idx, 52)]; month_name = calendar.month_name[i]; month_map[month_name] = month_cols
    # logging.info(f"Approx month map created. Example Jan: {month_map.get('January', [])}") # Reduce verbosity
    return month_map

def calculer_forecast_simulation_v2(df, all_semaine_columns, selected_month_names, sim_type, progression_pct=0, objectif_montant=0):
    """ Performs forecast simulation for SELECTED MONTHS based on corresponding N-1 data. """
    # --- (Code inchangé, utilise détection dynamique YYYYWW) ---
    try:
        if not isinstance(df, pd.DataFrame) or df.empty: st.warning("Aucune donnée pour simulation."); return None, 0.0
        if not all_semaine_columns: st.error("Aucune colonne de ventes hebdomadaires identifiée."); return None, 0.0
        if not selected_month_names: st.warning("Veuillez sélectionner au moins un mois."); return None, 0.0
        required_cols = ["Référence Article", "Désignation Article", "Conditionnement", "Tarif d'achat"]
        if not all(col in df.columns for col in required_cols): missing = [col for col in required_cols if col not in df.columns]; st.error(f"Colonnes manquantes simulation : {', '.join(missing)}"); return None, 0.0
        years_in_cols = set(); valid_week_cols_structure = []
        for col in all_semaine_columns:
            if isinstance(col, str) and len(col) >= 6 and col[:4].isdigit() and col[4:6].isdigit(): years_in_cols.add(int(col[:4])); valid_week_cols_structure.append(col)
        if not years_in_cols: st.error("Impossible de déterminer les années. Format attendu : 'YYYYWW'."); return None, 0.0
        if not valid_week_cols_structure: st.error("Aucune colonne au format 'YYYYWW' valide trouvée."); return None, 0.0
        current_year_n = max(years_in_cols); year_n_minus_1 = current_year_n - 1
        st.caption(f"(Année N={current_year_n}, Année N-1={year_n_minus_1})")
        n1_week_cols = [col for col in valid_week_cols_structure if col.startswith(str(year_n_minus_1))]; n1_week_cols.sort()
        logging.info(f"Forecast Sim v2: Identified {len(n1_week_cols)} N-1 cols for {year_n_minus_1}.")
        if len(n1_week_cols) < 52: st.error(f"Colonnes N-1 ({year_n_minus_1}) insuffisantes ({len(n1_week_cols)})."); return None, 0.0
        n1_week_cols_for_mapping = n1_week_cols[:52]
        df_sim = df[required_cols + ["Fournisseur"]].copy(); df_sim["Tarif d'achat"] = pd.to_numeric(df_sim["Tarif d'achat"], errors='coerce').fillna(0); df_sim["Conditionnement"] = pd.to_numeric(df_sim["Conditionnement"], errors='coerce').fillna(1).apply(lambda x: 1 if x<=0 else int(x))
        if not all(col in df.columns for col in n1_week_cols): missing_in_df = [col for col in n1_week_cols if col not in df.columns]; st.error(f"Erreur interne: Colonnes N-1 manquantes : {missing_in_df}"); return None, 0.0
        df_n1_sales = df[n1_week_cols].copy()
        for col in n1_week_cols: df_n1_sales[col] = pd.to_numeric(df_n1_sales[col], errors='coerce').fillna(0)
        month_col_map_n1 = approx_weeks_to_months(n1_week_cols_for_mapping); total_n1_sales_selected_months = pd.Series(0.0, index=df_sim.index); monthly_sales_n1_selected = {}
        for month in selected_month_names:
            if month in month_col_map_n1 and month_col_map_n1[month]:
                month_n1_cols_mapped = [col for col in month_col_map_n1[month] if col in df_n1_sales.columns]
                if month_n1_cols_mapped: sales_this_month = df_n1_sales[month_n1_cols_mapped].sum(axis=1); monthly_sales_n1_selected[month] = sales_this_month; total_n1_sales_selected_months += sales_this_month; df_sim[f"Ventes N-1 {month}"] = sales_this_month
                else: monthly_sales_n1_selected[month] = pd.Series(0.0, index=df_sim.index); df_sim[f"Ventes N-1 {month}"] = 0.0
            else: monthly_sales_n1_selected[month] = pd.Series(0.0, index=df_sim.index); df_sim[f"Ventes N-1 {month}"] = 0.0
        df_sim["Vts N-1 Tot (Mois Sel.)"] = total_n1_sales_selected_months; period_seasonality = {}; safe_total_n1_sales_selected = total_n1_sales_selected_months.replace(0, np.nan)
        for month in selected_month_names:
            if month in monthly_sales_n1_selected: period_seasonality[month] = (monthly_sales_n1_selected[month] / safe_total_n1_sales_selected).fillna(0)
            else: period_seasonality[month] = 0.0
        base_monthly_forecast_qty = {}
        if sim_type == 'Simple Progression':
            prog_factor = 1 + (progression_pct / 100.0); total_forecast_qty_selected_period = total_n1_sales_selected_months * prog_factor
            for month in selected_month_names: base_monthly_forecast_qty[month] = total_forecast_qty_selected_period * period_seasonality.get(month, 0.0)
        elif sim_type == 'Objectif Montant':
            if objectif_montant <= 0: st.error("Objectif > 0 requis."); return None, 0.0
            total_n1_sales_check = total_n1_sales_selected_months.sum()
            if total_n1_sales_check <= 0:
                st.warning("Ventes N-1 nulles. Répartition égale tentée."); num_sel_m = len(selected_month_names);
                if num_sel_m == 0: return None, 0.0
                amt_per_m = objectif_montant / num_sel_m
                for month in selected_month_names: base_monthly_forecast_qty[month] = np.divide(amt_per_m, df_sim["Tarif d'achat"], out=np.zeros_like(df_sim["Tarif d'achat"], dtype=float), where=df_sim["Tarif d'achat"]!=0)
            else:
                for month in selected_month_names:
                    target_amt_m = objectif_montant * period_seasonality.get(month, 0.0)
                    base_monthly_forecast_qty[month] = np.divide(target_amt_m, df_sim["Tarif d'achat"], out=np.zeros_like(df_sim["Tarif d'achat"], dtype=float), where=df_sim["Tarif d'achat"]!=0)
        else: st.error("Type sim non reconnu."); return None, 0.0
        df_result = df_sim[["Fournisseur", "Référence Article", "Désignation Article"]].copy(); df_result["Conditionnement"] = df_sim["Conditionnement"]; df_result["Tarif d'achat"] = df_sim["Tarif d'achat"]
        total_adjusted_qty_annual = pd.Series(0.0, index=df_result.index); all_month_cols = list(calendar.month_name)[1:]
        for i, month in enumerate(all_month_cols):
            month_qty_col = f"{month}"
            if month in selected_month_names and month in base_monthly_forecast_qty:
                 base_q = pd.to_numeric(base_monthly_forecast_qty[month], errors='coerce').fillna(0); cond = df_sim["Conditionnement"]
                 adjusted_qty = (np.ceil(np.divide(base_q, cond, out=np.zeros_like(base_q, dtype=float), where=cond!=0)) * cond).fillna(0).astype(int)
                 df_result[month_qty_col] = adjusted_qty; total_adjusted_qty_annual += adj_qty
            else: df_result[month_qty_col] = 0
        df_result["Total Annuel"] = total_adjusted_qty_annual
        id_cols_out = ["Référence Article", "Désignation Article"]; month_cols_out = all_month_cols; total_col_out = ["Total Annuel"]
        final_cols_ordered = id_cols_out + month_cols_out + total_col_out; final_cols_existing = [col for col in final_cols_ordered if col in df_result.columns]
        grand_total_amount = (df_result["Total Annuel"] * df_result["Tarif d'achat"]).sum()
        # Fusionner les infos N-1 si nécessaire
        df_result_final = pd.merge(df_result[final_cols_existing], df_sim[['Référence Article'] + [f"Ventes N-1 {m}" for m in selected_month_names if f"Ventes N-1 {m}" in df_sim.columns] + ["Vts N-1 Tot (Mois Sel.)"]].drop_duplicates(subset=['Référence Article']), on="Référence Article", how="left")
        # Réorganiser les colonnes
        final_cols_reordered = id_cols_out + ["Vts N-1 Tot (Mois Sel.)"] + [f"Ventes N-1 {m}" for m in selected_month_names if f"Ventes N-1 {m}" in df_result_final.columns] + month_cols_out + total_col_out
        final_cols_reordered_existing = [col for col in final_cols_reordered if col in df_result_final.columns]
        return df_result_final[final_cols_reordered_existing], grand_total_amount # Renvoyer DF ordonné et montant
    except Exception as e: st.error(f"Erreur simulation forecast v2 : {e}"); logging.exception("Error forecast sim v2:"); return None, 0.0

def sanitize_sheet_name(name):
    """ Removes invalid characters for Excel sheet names and truncates. """
    if not isinstance(name, str): name = str(name)
    sanitized = re.sub(r'[\[\]:*?/\\<>|"]', '_', name)
    if sanitized.startswith("'"): sanitized = "_" + sanitized[1:]
    if sanitized.endswith("'"): sanitized = sanitized[:-1] + "_"
    return sanitized[:31]

def render_supplier_checkboxes(tab_key_prefix, all_suppliers, default_select_all=False):
    """ Renders supplier checkboxes and returns selected list for the tab. """
    select_all_key = f"{tab_key_prefix}_select_all"
    supplier_keys = {supplier: f"{tab_key_prefix}_cb_{sanitized_supplier_key(supplier)}" for supplier in all_suppliers}
    if select_all_key not in st.session_state: st.session_state[select_all_key] = default_select_all
    for key in supplier_keys.values():
        if key not in st.session_state: st.session_state[key] = st.session_state[select_all_key]
    def toggle_all():
        is_checked = st.session_state[select_all_key]
        for key in supplier_keys.values(): st.session_state[key] = is_checked
    def check_individual():
        all_checked = all(st.session_state.get(key, False) for key in supplier_keys.values())
        if st.session_state[select_all_key] != all_checked: st.session_state[select_all_key] = all_checked
    with st.expander("👤 Sélectionner Fournisseurs", expanded=True):
        st.checkbox("Sélectionner / Désélectionner Tout", key=select_all_key, on_change=toggle_all, disabled=not bool(all_suppliers))
        st.markdown("---")
        selected_in_ui = []; num_cols = 4; cols = st.columns(num_cols); col_idx = 0
        for supplier in all_suppliers:
            key = supplier_keys[supplier]
            is_checked = cols[col_idx].checkbox(supplier, key=key, value=st.session_state.get(key, False), on_change=check_individual)
            if is_checked: selected_in_ui.append(supplier)
            col_idx = (col_idx + 1) % num_cols
    logging.debug(f"Checkboxes {tab_key_prefix}. Selected: {len(selected_in_ui)}")
    return selected_in_ui

def sanitized_supplier_key(supplier_name):
     """Creates a safe key for session state from supplier name."""
     s = re.sub(r'\W+', '_', supplier_name); s = re.sub(r'^_+|_+$', '', s); s = re.sub(r'_+', '_', s)
     return s if s else "invalid_key"


# --- Streamlit App Main Logic ---
st.set_page_config(page_title="Forecast & Rotation App", layout="wide")
st.title("📦 Application Prévision Commande & Analyse Rotation")

# --- File Upload ---
uploaded_file = st.file_uploader("📁 Charger le fichier Excel principal", type=["xlsx", "xls"], key="fileUploader")

# --- Initialize Session State ---
default_values = {'df_full': None, 'min_order_dict': {}, 'df_initial_filtered': pd.DataFrame(), 'semaine_columns': [], 'calculation_result_df': None, 'rotation_result_df': None, 'forecast_result_df': None, 'forecast_grand_total': 0.0, 'rotation_threshold_value': 1.0, 'show_all_rotation': True, 'forecast_selected_months': list(calendar.month_name)[1:], 'forecast_sim_type_index': 0, 'forecast_prog_pct': 5.0, 'forecast_target_amount': 10000.0, 'sel_fourn_calc_cmd': [], 'sel_fourn_calc_rot': [] }
for key, default_value in default_values.items():
    if key not in st.session_state: st.session_state[key] = default_value
# Initialize checkbox states dynamically if needed
tab_prefixes = ['tab1', 'tab2', 'tab4']
temp_fournisseurs_init = []
if 'df_initial_filtered' in st.session_state and not st.session_state.df_initial_filtered.empty and "Fournisseur" in st.session_state.df_initial_filtered.columns:
    temp_fournisseurs_init = sorted(st.session_state.df_initial_filtered["Fournisseur"].unique().tolist())
for prefix in tab_prefixes:
    select_all_key = f"{prefix}_select_all"
    if select_all_key not in st.session_state: st.session_state[select_all_key] = True # Default select all
    for supplier in temp_fournisseurs_init:
        key = f"{prefix}_cb_{sanitized_supplier_key(supplier)}"
        if key not in st.session_state: st.session_state[key] = st.session_state[select_all_key]

# --- Data Loading ---
if uploaded_file and st.session_state.df_full is None:
    logging.info(f"Processing new file: {uploaded_file.name}")
    keys_to_clear = [k for k in st.session_state if k != 'df_full']
    dynamic_keys = [k for k in st.session_state if k.startswith(('tab1_', 'tab2_', 'tab4_'))]
    keys_to_clear.extend(dynamic_keys)
    for key in keys_to_clear:
        if key in st.session_state: del st.session_state[key]
    for key, default_value in default_values.items():
         if key not in st.session_state: st.session_state[key] = default_value
    try:
        file_buffer = io.BytesIO(uploaded_file.getvalue()); st.info("Lecture 'Tableau final'...")
        df_full_temp = safe_read_excel(file_buffer, sheet_name="Tableau final", header=7)
        if df_full_temp is None: st.error("❌ Échec lecture 'Tableau final'."); st.stop()
        required_on_load = ["Stock", "Fournisseur", "AF_RefFourniss", "Tarif d'achat", "Conditionnement"]; missing = [c for c in required_on_load if c not in df_full_temp.columns]
        if missing: st.error(f"❌ Colonnes manquantes: {', '.join(missing)}"); st.stop()
        df_full_temp["Stock"] = pd.to_numeric(df_full_temp["Stock"], errors='coerce').fillna(0); df_full_temp["Tarif d'achat"] = pd.to_numeric(df_full_temp["Tarif d'achat"], errors='coerce').fillna(0); df_full_temp["Conditionnement"] = pd.to_numeric(df_full_temp["Conditionnement"], errors='coerce').fillna(1).apply(lambda x: 1 if x<=0 else int(x))
        st.session_state.df_full = df_full_temp; st.success("✅ 'Tableau final' lu.")
        st.info("Lecture 'Minimum de commande'...")
        df_min_temp = safe_read_excel(file_buffer, sheet_name="Minimum de commande"); min_dict_temp = {}
        if df_min_temp is not None:
            st.success("✅ 'Minimum de commande' lu."); sup_col = "Fournisseur"; min_col = "Minimum de Commande"; req_min_cols = [sup_col, min_col]
            if all(c in df_min_temp.columns for c in req_min_cols):
                try: df_min_temp[sup_col] = df_min_temp[sup_col].astype(str).str.strip(); df_min_temp[min_col] = pd.to_numeric(df_min_temp[min_col], errors='coerce'); min_dict_temp = df_min_temp.dropna(subset=[sup_col, min_col]).set_index(sup_col)[min_col].to_dict()
                except Exception as e: st.error(f"❌ Err traitement 'Min commande': {e}")
            else: st.warning(f"⚠️ Cols manquantes ({', '.join(req_min_cols)}) dans 'Min commande'.")
        st.session_state.min_order_dict = min_dict_temp
        df = st.session_state.df_full
        try:
            filter_cols = ["Fournisseur", "AF_RefFourniss"];
            if not all(c in df.columns for c in filter_cols): st.error(f"❌ Cols filtrage ({', '.join(filter_cols)}) manquantes."); st.stop()
            df_init_filtered = df[(df["Fournisseur"].notna()) & (df["Fournisseur"] != "") & (df["Fournisseur"] != "#FILTER") & (df["AF_RefFourniss"].notna()) & (df["AF_RefFourniss"] != "")].copy()
            st.session_state.df_initial_filtered = df_init_filtered
            start_col = 12; semaine_cols_temp = []
            if len(df.columns) > start_col:
                pot_w_cols = df.columns[start_col:].tolist(); exclude = ["Tarif d'achat", "Conditionnement", "Stock", "Total", "Stock à terme", "Ventes N-1", "Ventes 12 semaines identiques N-1", "Ventes 12 dernières semaines", "Quantité à commander", "Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article"]
                semaine_cols_temp = [c for c in pot_w_cols if c not in exclude and isinstance(c, str) and len(c)>=6 and c[:4].isdigit() and c[4:6].isdigit()]
                if not semaine_cols_temp: # Fallback
                     logging.warning("Format YYYYWW non détecté, fallback vers colonnes numériques.")
                     semaine_cols_temp = [c for c in pot_w_cols if c not in exclude and pd.api.types.is_numeric_dtype(df.get(c, pd.Series(dtype=float)).dtype)]
            st.session_state.semaine_columns = sorted(semaine_cols_temp) # Trier
            if not semaine_cols_temp: logging.warning("No valid week columns identified.")
            ess_num_cols = ["Stock", "Conditionnement", "Tarif d'achat"]; missing_ess = False
            for col in ess_num_cols:
                 if col in df_init_filtered.columns: df_init_filtered[col] = pd.to_numeric(df_init_filtered[col], errors='coerce').fillna(0)
                 elif not df_init_filtered.empty: st.error(f"Col essentielle '{col}' manquante."); missing_ess = True
            if missing_ess: st.stop()
            st.rerun()
        except KeyError as e: st.error(f"❌ Col filtrage '{e}' manquante."); st.stop()
        except Exception as e: st.error(f"❌ Err filtrage initial: {e}"); st.stop()
    except Exception as e: st.error(f"❌ Err chargement fichier: {e}"); logging.exception("File loading error:"); st.stop()


# --- Main App UI ---
if 'df_initial_filtered' in st.session_state and st.session_state.df_initial_filtered is not None:

    df_full = st.session_state.df_full
    df_base_filtered = st.session_state.get('df_initial_filtered', pd.DataFrame())
    fournisseurs_list_all = sorted(df_base_filtered["Fournisseur"].unique().tolist()) if not df_base_filtered.empty and "Fournisseur" in df_base_filtered.columns else []
    min_order_dict = st.session_state.min_order_dict
    semaine_columns = st.session_state.semaine_columns

    # --- NO SIDEBAR ---

    # --- Tabs ---
    tab1, tab2, tab3, tab4 = st.tabs(["Prévision Commande", "Analyse Rotation Stock", "Vérification Stock", "Simulation Forecast"])

    # ========================= TAB 1: Prévision Commande =========================
    with tab1:
        st.header("Prévision Quantités à Commander")
        selected_fournisseurs_tab1 = render_supplier_checkboxes("tab1", fournisseurs_list_all, default_select_all=True)
        if selected_fournisseurs_tab1:
            df_display_tab1 = df_base_filtered[df_base_filtered["Fournisseur"].isin(selected_fournisseurs_tab1)].copy()
            st.caption(f"{len(df_display_tab1)} articles pour {len(selected_fournisseurs_tab1)} fournisseur(s).")
        else: df_display_tab1 = pd.DataFrame(columns=df_base_filtered.columns)
        st.markdown("---")

        if not selected_fournisseurs_tab1: st.info("Sélectionnez fournisseur(s) ci-dessus.")
        elif df_display_tab1.empty: st.warning("Aucun article trouvé.")
        elif not semaine_columns: st.warning("Colonnes ventes manquantes.")
        else:
            st.markdown("#### Paramètres"); col1_cmd, col2_cmd = st.columns(2)
            with col1_cmd: duree_semaines_cmd = st.number_input(label="⏳ Durée couverture (sem.)", min_value=1, max_value=260, value=4, step=1, key="duree_cmd")
            with col2_cmd: montant_minimum_input_cmd = st.number_input(label="💶 Montant min global (€)", min_value=0.0, max_value=1e12, value=0.0, step=50.0, format="%.2f", key="montant_min_cmd")
            if st.button("🚀 Calculer Quantités", key="calc_cmd_btn"):
                with st.spinner("Calcul..."): result_cmd = calculer_quantite_a_commander(df_display_tab1, semaine_columns, montant_minimum_input_cmd, duree_semaines_cmd)
                if result_cmd:
                    st.success("✅ Calcul OK."); (q_calc, vN1_tot, v12N1_sim, v12l, mt_calc) = result_cmd; df_res_cmd = df_display_tab1.copy()
                    df_res_cmd["Qte Cmdée"] = q_calc; df_res_cmd["Vts N-1 Total (calc)"] = vN1_tot; df_res_cmd["Vts 12 N-1 Sim (calc)"] = v12N1_sim; df_res_cmd["Vts 12 Dern. (calc)"] = v12l
                    df_res_cmd["Tarif Ach."] = pd.to_numeric(df_res_cmd["Tarif d'achat"], errors='coerce').fillna(0); df_res_cmd["Total Cmd"] = df_res_cmd["Tarif Ach."] * df_res_cmd["Qte Cmdée"]; df_res_cmd["Stock Terme"] = df_res_cmd["Stock"] + df_res_cmd["Qte Cmdée"]
                    st.session_state.calc_res_df = df_res_cmd; st.session_state.mt_calc = mt_calc; st.session_state.sel_fourn_calc_cmd = selected_fournisseurs_tab1
                    st.rerun()
                else: st.error("❌ Calcul échoué.");
            if 'calc_res_df' in st.session_state and st.session_state.calc_res_df is not None:
                if st.session_state.sel_fourn_calc_cmd == selected_fournisseurs_tab1:
                    st.markdown("---"); st.markdown("#### Résultats Commande"); df_cmd_disp = st.session_state.calc_res_df; mt_cmd_disp = st.session_state.mt_calc; sup_cmd_disp = st.session_state.sel_fourn_calc_cmd
                    st.metric(label="💰 Montant Total", value=f"{mt_cmd_disp:,.2f} €")
                    # Min Warning (Corrected Indentation)
                    if len(sup_cmd_disp) == 1:
                        sup_cmd = sup_cmd_disp[0]
                        if sup_cmd in min_order_dict:
                            req_min = min_order_dict.get(sup_cmd, 0)
                            if "Total Cmd" in df_cmd_disp.columns:
                                act_tot = df_cmd_disp["Total Cmd"].sum();
                                if req_min > 0 and act_tot < req_min:
                                    diff = req_min - act_tot; st.warning(f"⚠️ Min Non Atteint ({sup_cmd})\nMontant: **{act_tot:,.2f}€** | Requis: **{req_min:,.2f}€** (Manque: {diff:,.2f}€)")
                            else: logging.warning("Col 'Total Cmd' absente.")
                    # Display Table
                    cols_req = ["Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article", "Stock"]; cols_base = cols_req + ["Vts N-1 Total (calc)", "Vts 12 N-1 Sim (calc)", "Vts 12 Dern. (calc)", "Conditionnement", "Qte Cmdée", "Stock Terme", "Tarif Ach.", "Total Cmd"]
                    cols_disp = [c for c in cols_base if c in df_cmd_disp.columns];
                    if any(c not in df_cmd_disp.columns for c in cols_req): st.error("❌ Cols manquantes affichage.")
                    else: st.dataframe(df_cmd_disp[cols_disp].style.format({"Tarif Ach.": "{:,.2f}€", "Total Cmd": "{:,.2f}€", "Vts N-1 Total (calc)": "{:,.0f}", "Vts 12 N-1 Sim (calc)": "{:,.0f}", "Vts 12 Dern. (calc)": "{:,.0f}", "Stock": "{:,.0f}", "Conditionnement": "{:,.0f}", "Qte Cmdée": "{:,.0f}", "Stock Terme": "{:,.0f}"}, na_rep="-", thousands=","))
                    # Export Logic
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
                                     for sup_exp in sup_cmd_disp: # Use suppliers for whom calc was run
                                         df_sup_exp = df_exp_cmd[df_exp_cmd["Fournisseur"] == sup_exp].copy();
                                         if not df_sup_exp.empty:
                                             df_sh_data = df_sup_exp[export_cols_cmd].copy(); n_rows = len(df_sh_data); tot_v = df_sh_data[tot_c].sum(); req_m = min_order_dict.get(sup_exp, 0); min_f = f"{req_m:,.2f}€" if req_m > 0 else "N/A"
                                             lbl_c = "Désignation Article" if "Désignation Article" in export_cols_cmd else export_cols_cmd[1]; tot_r = {c: "" for c in export_cols_cmd}; tot_r[lbl_c] = "TOTAL"; tot_r[tot_c] = tot_v; min_r = {c: "" for c in export_cols_cmd}; min_r[lbl_c] = "Min Requis"; min_r[tot_c] = min_f
                                             df_sh = pd.concat([df_sh_data, pd.DataFrame([tot_r]), pd.DataFrame([min_r])], ignore_index=True); s_name = sanitize_sheet_name(sup_exp)
                                             try: # Write sheet and formulas
                                                 df_sh.to_excel(writer_cmd, sheet_name=s_name, index=False); ws = writer_cmd.sheets[s_name];
                                                 for r in range(2, n_rows + 2): form = f"={qty_l}{r}*{price_l}{r}"; cell = ws[f"{tot_l}{r}"]; cell.value = form; cell.number_format = '#,##0.00€'
                                                 total_formula_row_cmd = n_rows + 2
                                                 if n_rows > 0: sum_form = f"=SUM({tot_l}2:{tot_l}{n_rows + 1})"; s_cell = ws[f"{tot_l}{total_formula_row_cmd}"]; s_cell.value = sum_form; s_cell.number_format = '#,##0.00€'
                                                 sheets_cr_cmd += 1
                                             except Exception as we: logging.exception(f"Err sheet {s_name}:{we}")
                                 else: st.error("Export CMD: Erreur cols formules.")
                         except Exception as e_w: logging.exception(f"Err writer:{e_w}")
                         if sheets_cr_cmd > 0:
                              out_cmd.seek(0); fname = f"commande_{'multi' if len(sup_cmd_disp)>1 else sanitize_sheet_name(sup_cmd_disp[0])}_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx"
                              st.download_button(f"📥 Télécharger ({sheets_cr_cmd})", out_cmd, fname, key="dl_cmd_btn")
                    else: st.info("Aucune qté > 0 à exporter.")
                else: st.info("Résultats précédents invalidés. Relancez calcul.")


    # ====================== TAB 2: Analyse Rotation Stock ======================
    with tab2:
        st.header("Analyse Rotation Stocks")
        selected_fournisseurs_tab2 = render_supplier_checkboxes("tab2", fournisseurs_list_all, default_select_all=True)
        if selected_fournisseurs_tab2:
            df_display_tab2 = df_base_filtered[df_base_filtered["Fournisseur"].isin(selected_fournisseurs_tab2)].copy()
            st.caption(f"{len(df_display_tab2)} articles pour {len(selected_fournisseurs_tab2)} fournisseur(s).")
        else: df_display_tab2 = pd.DataFrame(columns=df_base_filtered.columns)
        st.markdown("---")
        if not selected_fournisseurs_tab2: st.info("Sélectionnez fournisseur(s).")
        elif df_display_tab2.empty: st.warning("Aucun article trouvé.")
        elif not semaine_columns: st.warning("Colonnes ventes manquantes.")
        else:
            st.markdown("#### Paramètres"); col1_r, col2_r = st.columns(2)
            with col1_r: period_opts = {"12 sem.": 12, "52 sem.": 52, "Total": 0}; sel_p_lbl = st.selectbox("📅 Période:", period_opts.keys(), key="rot_p_sel"); sel_p_w = period_opts[sel_p_lbl]
            with col2_r: st.markdown("##### Options Affichage"); show_all = st.checkbox("Afficher tout", value=st.session_state.show_all_rotation, key="show_all_rot_cb"); st.session_state.show_all_rotation = show_all; rot_thr = st.number_input("... ou ventes mens. <", 0.0, value=st.session_state.rotation_threshold_value, step=0.1, format="%.1f", key="rot_thr_in", disabled=show_all)
            if not show_all: st.session_state.rotation_threshold_value = rot_thr
            if st.button("🔄 Analyser Rotation", key="analyze_rot_btn"):
                 with st.spinner("Analyse..."): df_rot_res = calculer_rotation_stock(df_display_tab2, semaine_columns, sel_p_w)
                 if df_rot_res is not None: st.success("✅ Analyse terminée."); st.session_state.rot_res_df = df_rot_res; st.session_state.rot_p_lbl = sel_p_lbl; st.session_state.sel_fourn_calc_rot = selected_fournisseurs_tab2; st.rerun()
                 else: st.error("❌ Analyse échouée.");
            if 'rot_res_df' in st.session_state and st.session_state.rot_res_df is not None:
                 if st.session_state.sel_fourn_calc_rot == selected_fournisseurs_tab2:
                    st.markdown("---"); st.markdown(f"#### Résultats Rotation ({st.session_state.get('rot_p_lbl', '')})"); df_rot_orig = st.session_state.rot_res_df; thr_disp = st.session_state.rotation_threshold_value; show_all_f = st.session_state.show_all_rotation
                    m_sales_col = "Ventes Moy Mensuel (Période)"; can_filt = False; df_rot_disp = pd.DataFrame()
                    if m_sales_col in df_rot_orig.columns: m_sales_ser = pd.to_numeric(df_rot_orig[m_sales_col], errors='coerce').fillna(0); can_filt = True
                    else: st.warning(f"Col '{m_sales_col}' non trouvée.")
                    # CORRECTED Filter block structure
                    if show_all_f:
                        df_rot_disp = df_rot_orig.copy()
                        st.caption(f"Affichage {len(df_rot_disp)} articles.")
                    elif can_filt:
                        try:
                             df_rot_disp = df_rot_orig[m_sales_ser < thr_disp].copy()
                             st.caption(f"Filtre: Ventes < {thr_disp:.1f}/mois. {len(df_rot_disp)} / {len(df_rot_orig)} articles.")
                        except Exception as ef:
                            st.error(f"Err filtre: {ef}")
                            df_rot_disp = df_rot_orig.copy() # Fallback
                    else: # Cannot filter
                         df_rot_disp = df_rot_orig.copy()
                         if not show_all_f: st.info("Affichage tous résultats car filtrage impossible.")

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
                         out_r.seek(0); sups_exp = selected_fournisseurs_tab2 # Use this tab's selection
                         f_rot = f"{f_base}_{'multi' if len(sups_exp)>1 else sanitize_sheet_name(sups_exp[0] if sups_exp else 'NA')}_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx"
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
        st.header("Simulation Forecast Annuel")
        selected_fournisseurs_tab4 = render_supplier_checkboxes("tab4", fournisseurs_list_all, default_select_all=True)
        if selected_fournisseurs_tab4:
            df_display_tab4 = df_base_filtered[df_base_filtered["Fournisseur"].isin(selected_fournisseurs_tab4)].copy()
            st.caption(f"{len(df_display_tab4)} articles pour {len(selected_fournisseurs_tab4)} fournisseur(s).")
        else: df_display_tab4 = pd.DataFrame(columns=df_base_filtered.columns)
        st.markdown("---")
        st.caption("Simulation basée sur N-1 (colonnes identifiées par année YYYY)."); st.warning("🚨 **Approximation Importante:** Saisonnalité mensuelle basée sur découpage approx. des 52 sem. N-1.")

        if not selected_fournisseurs_tab4: st.info("Veuillez sélectionner un ou plusieurs fournisseurs ci-dessus.")
        elif df_display_tab4.empty: st.warning("Aucun article trouvé.")
        # CORRECTED Check: Need at least 52 columns overall for N-1 potential
        elif not semaine_columns or len(semaine_columns) < 52:
            st.warning("Données historiques insuffisantes (moins de 52 colonnes ventes identifiées).")
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
                    with st.spinner("Simulation..."):
                        # Utiliser la fonction V2 (avec détection dynamique N-1)
                        df_fcst_res, grand_total = calculer_forecast_simulation_v2(df_display_tab4, semaine_columns, sel_months_fcst, sim_t, prog_use, obj_use)
                    if df_fcst_res is not None: st.success("✅ Simulation terminée."); st.session_state.forecast_result_df = df_fcst_res; st.session_state.forecast_grand_total = grand_total; st.session_state.forecast_params = {'suppliers': selected_fournisseurs_tab4, 'months': sel_months_fcst, 'type': sim_t, 'prog': prog_use, 'obj': obj_use}; st.rerun() # Store tab-specific selection
                    else: st.error("❌ Simulation échouée.");
            if 'forecast_result_df' in st.session_state and st.session_state.forecast_result_df is not None:
                 # Comparer avec la sélection de cet onglet
                 current_params_disp = {'suppliers': selected_fournisseurs_tab4, 'months': sel_months_fcst, 'type': sim_t, 'prog': st.session_state.get('forecast_prog_pct', 5.0) if sim_t=='Simple Progression' else 0, 'obj': st.session_state.get('forecast_target_amount', 10000.0) if sim_t=='Objectif Montant' else 0}
                 if st.session_state.get('forecast_params') == current_params_disp:
                    st.markdown("---"); st.markdown("#### Résultats Simulation")
                    df_fcst_disp = st.session_state.forecast_result_df; grand_total_disp = st.session_state.forecast_grand_total

                    # Utiliser les colonnes retournées par la fonction v2
                    fcst_disp_fin = df_fcst_disp.columns.tolist() # Colonnes déjà ordonnées

                    if df_fcst_disp.empty: st.info("Aucun résultat.")
                    elif not fcst_disp_fin: st.error("Erreur: Colonnes de résultats de prévision manquantes.")
                    else:
                        # Définir les formateurs pour les colonnes attendues
                        fcst_fmters_final = {}
                        id_cols = ["Référence Article", "Désignation Article"]
                        for col in df_fcst_disp.columns:
                             if col not in id_cols: # Formater toutes les autres comme nombres entiers (Qté ou Ventes N-1)
                                 fcst_fmters_final[col] = "{:,.0f}"
                             # Ajouter format spécifique pour tarif si affiché (normalement non)
                             # elif col == "Tarif d'achat": fcst_fmters_final[col] = "{:,.2f}€"

                        try: st.dataframe(df_fcst_disp.style.format(fcst_fmters_final, na_rep="-", thousands=","))
                        except Exception as e_fmt: st.error(f"Erreur formatage affichage: {e_fmt}"); st.dataframe(df_fcst_disp)

                        st.metric(label="Montant Total Général Prévisionnel (€)", value=f"{grand_total_disp:,.2f} €", help="Calculé en multipliant le 'Total Annuel' prévu (quantité) par le 'Tarif d'achat' actuel de chaque article.")

                        # Export Forecast Results
                        st.markdown("#### Export Simulation"); out_f = io.BytesIO(); df_exp_f = df_fcst_disp.copy() # Exporter toutes les colonnes retournées
                        # Ajouter le total général en bas pour l'export
                        if not df_exp_f.empty:
                             try:
                                 total_row_data = {}; label_col_export = "Désignation Article" if "Désignation Article" in df_exp_f.columns else df_exp_f.columns[1]
                                 total_col_export = "Total Annuel"
                                 for col in df_exp_f.columns: total_row_data[col] = ''
                                 total_row_data[label_col_export] = 'TOTAL GÉNÉRAL';
                                 if total_col_export in df_exp_f.columns: total_row_data[total_col_export] = df_exp_f[total_col_export].sum()
                                 total_row_fcst = pd.DataFrame([total_row_data]); df_exp_f = pd.concat([df_exp_f, total_row_fcst], ignore_index=True)
                             except Exception as e_total: logging.error(f"Err ajout total export forecast: {e_total}")

                        try:
                            with pd.ExcelWriter(out_f, engine="openpyxl") as w_f: df_exp_f.to_excel(w_f, sheet_name=f"Forecast_{sim_t.replace(' ','_')}", index=False)
                            out_f.seek(0); fb = f"forecast_{sim_t.replace(' ','_').lower()}"; sups_f = selected_fournisseurs_tab4 # Utiliser sélection onglet
                            f_fcst = f"{fb}_{'multi' if len(sups_f)>1 else sanitize_sheet_name(sups_f[0] if sups_f else 'NA')}_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx"; st.download_button("📥 Télécharger Simulation", out_f, f_fcst, key="dl_fcst_btn")
                        except Exception as eef: st.error(f"Err export forecast: {eef}")
                 else: st.info("Résultats simulation invalidés. Relancez.")


# --- App footer/initial message ---
elif not uploaded_file:
    st.info("👋 Bienvenue ! Chargez votre fichier Excel pour commencer.")
    if st.button("🔄 Réinitialiser l'application"):
         keys_to_clear = list(st.session_state.keys())
         dynamic_keys = [k for k in st.session_state if k.startswith(('tab1_', 'tab2_', 'tab4_'))]
         keys_to_clear.extend(dynamic_keys)
         for key in keys_to_clear:
             if key in st.session_state: del st.session_state[key]
         st.rerun() # Use st.rerun()
