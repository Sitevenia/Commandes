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

# ==============================================================================
# --- NEW: Forecast Simulation Calculation Function --- (Copied from previous answer)
# ==============================================================================
def approx_weeks_to_months(week_columns_52):
    """Approximates month mapping for 52 consecutive week columns."""
    month_map = {}
    weeks_per_month_approx = 52 / 12
    for i in range(1, 13):
        start_idx = int(round((i-1) * weeks_per_month_approx))
        end_idx = int(round(i * weeks_per_month_approx))
        # Ensure indices are within bounds and slice correctly
        month_cols = week_columns_52[start_idx : min(end_idx, 52)]
        month_name = calendar.month_name[i]
        month_map[month_name] = month_cols
    logging.info(f"Approx month map created. Example Jan: {month_map.get('January', [])}")
    return month_map

def calculer_forecast_simulation(df, all_semaine_columns, selected_months, sim_type, progression_pct=0, objectif_montant=0):
    """ Performs forecast simulation based on N-1 data (last 52 weeks assumed). """
    try:
        if not isinstance(df, pd.DataFrame) or df.empty: st.warning("Aucune donnée pour simulation."); return None
        if len(all_semaine_columns) < 52: st.error("Données historiques insuffisantes (< 52 semaines)."); return None
        required_cols = ["Référence Article", "Désignation Article", "Conditionnement", "Tarif d'achat"]
        if not all(col in df.columns for col in required_cols): missing = [col for col in required_cols if col not in df.columns]; st.error(f"Colonnes manquantes simulation : {', '.join(missing)}"); return None

        df_sim = df[required_cols + ["Fournisseur"]].copy()
        n1_week_cols = all_semaine_columns[-52:]; df_n1_sales = df[n1_week_cols].copy()
        for col in n1_week_cols: df_n1_sales[col] = pd.to_numeric(df_n1_sales[col], errors='coerce').fillna(0)
        df_sim["Ventes Totales N-1"] = df_n1_sales.sum(axis=1)

        month_col_map = approx_weeks_to_months(n1_week_cols); monthly_sales_n1 = {}; seasonality = {}
        total_n1_annual_sales = df_sim["Ventes Totales N-1"]
        for month, week_cols in month_col_map.items():
            if not week_cols: continue
            monthly_sales_n1[month] = df_n1_sales[week_cols].sum(axis=1)
            df_sim[f"Ventes N-1 {month}"] = monthly_sales_n1[month]
            seasonality[month] = np.divide(monthly_sales_n1[month], total_n1_annual_sales, out=np.zeros_like(total_n1_annual_sales, dtype=float), where=total_n1_annual_sales!=0)

        total_forecast_qty = pd.Series(0.0, index=df_sim.index); base_monthly_forecast_qty = {}
        if sim_type == 'Simple Progression':
            prog_factor = 1 + (progression_pct / 100.0); total_forecast_qty = total_n1_annual_sales * prog_factor
            for month in selected_months: base_monthly_forecast_qty[month] = total_forecast_qty * seasonality.get(month, 0.0)
        elif sim_type == 'Objectif Montant':
            if objectif_montant <= 0: st.error("Objectif montant > 0 requis."); return None
            df_sim["Tarif d'achat"] = pd.to_numeric(df_sim["Tarif d'achat"], errors='coerce').fillna(0)
            initial_total_amount_n1_based = (total_n1_annual_sales * df_sim["Tarif d'achat"]).sum()
            if initial_total_amount_n1_based <= 0:
                 st.warning("Calcul facteur échelle impossible (ventes ou tarifs N-1 nuls). Approche alternative.")
                 total_price_weight = (df_sim["Tarif d'achat"] * df_sim["Tarif d'achat"].ne(0)).sum()
                 if total_price_weight == 0 : st.error("Tarifs nuls/manquants."); return None
                 for month in selected_months:
                     if month in seasonality:
                         n1_month_value = (monthly_sales_n1[month] * df_sim["Tarif d'achat"]); total_n1_value = (total_n1_annual_sales * df_sim["Tarif d'achat"])
                         value_seasonality = np.divide(n1_month_value, total_n1_value, out=np.zeros_like(n1_month_value), where=total_n1_value!=0)
                         target_month_amount = objectif_montant * value_seasonality
                         base_monthly_forecast_qty[month] = np.divide(target_month_amount, df_sim["Tarif d'achat"], out=np.zeros_like(target_month_amount), where=df_sim["Tarif d'achat"]!=0)
                     else: base_monthly_forecast_qty[month] = pd.Series(0.0, index=df_sim.index)
            else:
                scaling_factor = objectif_montant / initial_total_amount_n1_based; total_forecast_qty = total_n1_annual_sales * scaling_factor
                for month in selected_months: base_monthly_forecast_qty[month] = total_forecast_qty * seasonality.get(month, 0.0)
        else: st.error("Type simulation non reconnu."); return None

        df_sim["Conditionnement"] = pd.to_numeric(df_sim["Conditionnement"], errors='coerce').fillna(1).apply(lambda x: 1 if x<=0 else int(x)) # Ensure int Cond >= 1
        total_adjusted_qty = pd.Series(0.0, index=df_sim.index); total_final_amount = pd.Series(0.0, index=df_sim.index)
        for month in selected_months:
            month_qty_col = f"Qté Prév. {month}"; month_amt_col = f"Montant Prév. {month} (€)"
            if month in base_monthly_forecast_qty:
                 base_qty = pd.to_numeric(base_monthly_forecast_qty[month], errors='coerce').fillna(0)
                 cond = df_sim["Conditionnement"]
                 adjusted_qty = (np.ceil(np.divide(base_qty, cond, out=np.zeros_like(base_qty), where=cond!=0)) * cond).fillna(0).astype(int)
                 df_sim[month_qty_col] = adjusted_qty
                 df_sim[month_amt_col] = adjusted_qty * df_sim["Tarif d'achat"]
                 total_adjusted_qty += adjusted_qty; total_final_amount += df_sim[month_amt_col]
            else: df_sim[month_qty_col] = 0; df_sim[month_amt_col] = 0.0

        df_sim["Qté Totale Prév."] = total_adjusted_qty; df_sim["Montant Total Prév. (€)"] = total_final_amount

        id_cols = ["Fournisseur", "Référence Article", "Désignation Article", "Conditionnement", "Tarif d'achat"]
        n1_cols = [f"Ventes N-1 {m}" for m in selected_months if f"Ventes N-1 {m}" in df_sim.columns]
        qty_cols = [f"Qté Prév. {m}" for m in selected_months]; amt_cols = [f"Montant Prév. {m} (€)" for m in selected_months]
        total_cols = ["Ventes Totales N-1", "Qté Totale Prév.", "Montant Total Prév. (€)"]
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
if 'forecast_result_df' not in st.session_state: st.session_state.forecast_result_df = None # For new tab
if 'selected_fournisseurs_session' not in st.session_state: st.session_state.selected_fournisseurs_session = []
if 'rotation_threshold_value' not in st.session_state: st.session_state.rotation_threshold_value = 1.0 # Default threshold
if 'show_all_rotation' not in st.session_state: st.session_state.show_all_rotation = True # Default to showing all

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
        required_on_load = ["Stock", "Fournisseur", "AF_RefFourniss", "Tarif d'achat", "Conditionnement"]
        if not all(col in df_full_temp.columns for col in required_on_load): missing_on_load = [col for col in required_on_load if col not in df_full_temp.columns]; st.error(f"❌ Colonnes manquantes: {', '.join(missing_on_load)}"); st.stop()
        df_full_temp["Stock"] = pd.to_numeric(df_full_temp["Stock"], errors='coerce').fillna(0)
        df_full_temp["Tarif d'achat"] = pd.to_numeric(df_full_temp["Tarif d'achat"], errors='coerce').fillna(0)
        df_full_temp["Conditionnement"] = pd.to_numeric(df_full_temp["Conditionnement"], errors='coerce').fillna(1).apply(lambda x: 1 if x<=0 else int(x))
        st.session_state.df_full = df_full_temp
        st.success("✅ Onglet 'Tableau final' lu.")

        st.info("Lecture onglet 'Minimum de commande'...")
        df_min_commande_temp = safe_read_excel(file_buffer, sheet_name="Minimum de commande")
        min_order_dict_temp = {}
        if df_min_commande_temp is not None:
            st.success("✅ Onglet 'Minimum de commande' lu.")
            supplier_col_min = "Fournisseur"; min_amount_col = "Minimum de Commande"; required_min_cols = [supplier_col_min, min_amount_col]
            if all(col in df_min_commande_temp.columns for col in required_min_cols):
                try:
                    df_min_commande_temp[supplier_col_min] = df_min_commande_temp[supplier_col_min].astype(str).str.strip()
                    df_min_commande_temp[min_amount_col] = pd.to_numeric(df_min_commande_temp[min_amount_col], errors='coerce')
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
                potential_week_cols = df.columns[start_col_index:].tolist()
                exclude_cols = ["Tarif d'achat", "Conditionnement", "Stock", "Total", "Stock à terme", "Ventes N-1", "Ventes 12 semaines identiques N-1", "Ventes 12 dernières semaines", "Quantité à commander", "Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article"]
                semaine_cols_temp = [col for col in potential_week_cols if col not in exclude_cols and pd.api.types.is_numeric_dtype(df.get(col, pd.Series(dtype=float)).dtype)]
            st.session_state.semaine_columns = semaine_cols_temp
            if not semaine_cols_temp: logging.warning("No week columns identified after initial processing.")

            essential_numeric_cols = ["Stock", "Conditionnement", "Tarif d'achat"]; missing_essential = False
            for col in essential_numeric_cols:
                 if col in df_init_filtered.columns: df_init_filtered[col] = pd.to_numeric(df_init_filtered[col], errors='coerce').fillna(0)
                 elif not df_init_filtered.empty: st.error(f"Colonne essentielle '{col}' manquante."); missing_essential = True
            if missing_essential: st.stop()
            st.rerun() # Rerun to apply session state

        except KeyError as e_filter: st.error(f"❌ Colonne filtrage '{e_filter}' manquante."); st.stop()
        except Exception as e_filter_other: st.error(f"❌ Erreur filtrage initial : {e_filter_other}"); st.stop()
    except Exception as e_load: st.error(f"❌ Erreur lecture fichier : {e_load}"); logging.exception("File loading error:"); st.stop()


# --- Main App UI (Tabs) ---
if 'df_initial_filtered' in st.session_state and st.session_state.df_initial_filtered is not None:

    df_full = st.session_state.df_full
    df_base_filtered = st.session_state.get('df_initial_filtered', pd.DataFrame())
    fournisseurs_list = sorted(df_base_filtered["Fournisseur"].unique().tolist()) if not df_base_filtered.empty and "Fournisseur" in df_base_filtered.columns else []
    min_order_dict = st.session_state.min_order_dict
    semaine_columns = st.session_state.semaine_columns

    st.sidebar.header("Filtres (pour Prévision & Rotation)")
    selected_fournisseurs = st.sidebar.multiselect(
        "👤 Fournisseur(s)", options=fournisseurs_list,
        default=st.session_state.selected_fournisseurs_session,
        key="supplier_select_sidebar", disabled=not bool(fournisseurs_list),
        help="Filtre les données utilisées dans les onglets 'Prévision Commande' et 'Analyse Rotation Stock'."
    )
    st.session_state.selected_fournisseurs_session = selected_fournisseurs

    if selected_fournisseurs:
        df_display_filtered = df_base_filtered[df_base_filtered["Fournisseur"].isin(selected_fournisseurs)].copy()
        if df_display_filtered.empty and fournisseurs_list: st.sidebar.warning("Aucun article trouvé pour cette sélection.")
        elif not df_display_filtered.empty: st.sidebar.info(f"{len(df_display_filtered)} articles sélectionnés pour analyse.")
    else:
        df_display_filtered = df_base_filtered.copy()
        if not selected_fournisseurs and fournisseurs_list: st.sidebar.info("Affichage pour tous les fournisseurs filtrés initialement.")


    # --- Create Tabs --- ## !! CORRECTED LINE !! ##
    tab1, tab2, tab3, tab4 = st.tabs(["Prévision Commande", "Analyse Rotation Stock", "Vérification Stock", "Simulation Forecast"])

    # ========================= TAB 1: Prévision Commande =========================
    with tab1:
        st.header("Prévision des Quantités à Commander")
        st.caption("Utilise les fournisseurs sélectionnés dans la barre latérale.")

        if df_display_filtered.empty:
             if selected_fournisseurs: st.warning("Aucun article trouvé pour le(s) fournisseur(s) sélectionné(s).")
             else: st.info("Veuillez sélectionner au moins un fournisseur.")
        elif not semaine_columns: st.warning("Impossible de calculer: Aucune colonne de ventes valide identifiée.")
        else:
            st.markdown("#### Paramètres de Calcul")
            col1_cmd, col2_cmd = st.columns(2)
            with col1_cmd: duree_semaines_cmd = st.number_input(label="⏳ Durée couverture (semaines)", min_value=1, max_value=260, value=4, step=1, key="duree_cmd")
            with col2_cmd: montant_minimum_input_cmd = st.number_input(label="💶 Montant minimum global (€)", min_value=0.0, max_value=1e12, value=0.0, step=50.0, format="%.2f", key="montant_min_cmd")

            if st.button("🚀 Calculer les Quantités", key="calculate_button_cmd"):
                with st.spinner("Calcul en cours..."): result_cmd = calculer_quantite_a_commander(df_display_filtered, semaine_columns, montant_minimum_input_cmd, duree_semaines_cmd)
                if result_cmd is not None:
                    st.success("✅ Calculs terminés.")
                    (quantite_calc, vN1, v12N1, v12last, mt_calc) = result_cmd; df_result_cmd = df_display_filtered.copy()
                    df_result_cmd.loc[:, "Quantité à commander"] = quantite_calc; df_result_cmd.loc[:, "Ventes N-1"] = vN1; df_result_cmd.loc[:, "Ventes 12 semaines identiques N-1"] = v12N1; df_result_cmd.loc[:, "Ventes 12 dernières semaines"] = v12last
                    df_result_cmd.loc[:, "Tarif d'achat"] = pd.to_numeric(df_result_cmd["Tarif d'achat"], errors='coerce').fillna(0)
                    df_result_cmd.loc[:, "Total"] = df_result_cmd["Tarif d'achat"] * df_result_cmd["Quantité à commander"]; df_result_cmd.loc[:, "Stock à terme"] = df_result_cmd["Stock"] + df_result_cmd["Quantité à commander"]
                    st.session_state.calculation_result_df = df_result_cmd; st.session_state.montant_total_calc = mt_calc; st.session_state.selected_fournisseurs_calc_cmd = selected_fournisseurs
                    st.rerun()
                else:
                     st.error("❌ Le calcul des quantités a échoué.")
                     if 'calculation_result_df' in st.session_state: del st.session_state.calculation_result_df

            # Display Command Results
            if 'calculation_result_df' in st.session_state and st.session_state.calculation_result_df is not None:
                if st.session_state.selected_fournisseurs_calc_cmd == selected_fournisseurs:
                    st.markdown("---"); st.markdown("#### Résultats du Calcul de Commande")
                    df_results_cmd_display = st.session_state.calculation_result_df; montant_total_cmd_display = st.session_state.montant_total_calc; suppliers_cmd_displayed = st.session_state.selected_fournisseurs_calc_cmd
                    st.metric(label="💰 Montant total GLOBAL calculé", value=f"{montant_total_cmd_display:,.2f} €")
                    if len(suppliers_cmd_displayed) == 1: # Min Warning
                        supplier_cmd = suppliers_cmd_displayed[0]
                        if supplier_cmd in min_order_dict:
                            req_min_cmd = min_order_dict[supplier_cmd]; actual_total_cmd = df_results_cmd_display["Total"].sum()
                            if req_min_cmd > 0 and actual_total_cmd < req_min_cmd: diff_cmd = req_min_cmd - actual_total_cmd; st.warning(f"⚠️ **Min Non Atteint ({supplier_cmd})**\nMontant: **{actual_total_cmd:,.2f} €** | Requis: **{req_min_cmd:,.2f} €** (Manque: {diff_cmd:,.2f} €)\n➡️ Suggestion: Modifiez 'Montant min global (€)' et relancez.")
                    # Display Table
                    cmd_required_cols = ["Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article", "Stock"]; cmd_display_cols_base = cmd_required_cols + ["Ventes N-1", "Ventes 12 semaines identiques N-1", "Ventes 12 dernières semaines", "Conditionnement", "Quantité à commander", "Stock à terme", "Tarif d'achat", "Total"]
                    cmd_display_cols = [col for col in cmd_display_cols_base if col in df_results_cmd_display.columns]
                    if any(col not in df_results_cmd_display.columns for col in cmd_required_cols): st.error("❌ Colonnes manquantes affichage cmd.")
                    else: st.dataframe(df_results_cmd_display[cmd_display_cols].style.format({"Tarif d'achat": "{:,.2f}€", "Total": "{:,.2f}€", "Ventes N-1": "{:,.0f}", "Ventes 12 semaines identiques N-1": "{:,.0f}", "Ventes 12 dernières semaines": "{:,.0f}", "Stock": "{:,.0f}", "Conditionnement": "{:,.0f}", "Quantité à commander": "{:,.0f}", "Stock à terme": "{:,.0f}"}, na_rep="-", thousands=","))
                    # Export Logic
                    st.markdown("#### Exportation de la Commande Calculée")
                    df_export_cmd = df_results_cmd_display[df_results_cmd_display["Quantité à commander"] > 0].copy()
                    if not df_export_cmd.empty:
                         output_cmd = io.BytesIO(); sheets_created_cmd = 0
                         # --- (Multi-sheet export logic with formulas - unchanged) ---
                         try:
                             with pd.ExcelWriter(output_cmd, engine="openpyxl") as writer_cmd:
                                 qty_col_name_cmd = "Quantité à commander"; price_col_name_cmd = "Tarif d'achat"; total_col_name_cmd = "Total"; export_columns_cmd = [col for col in cmd_display_cols if col != 'Fournisseur']; formula_ready_cmd = False
                                 if all(c in export_columns_cmd for c in [qty_col_name_cmd, price_col_name_cmd, total_col_name_cmd]):
                                     try:
                                         qty_col_idx_cmd = export_columns_cmd.index(qty_col_name_cmd); price_col_idx_cmd = export_columns_cmd.index(price_col_name_cmd); total_col_idx_cmd = export_columns_cmd.index(total_col_name_cmd)
                                         qty_col_letter_cmd = get_column_letter(qty_col_idx_cmd + 1); price_col_letter_cmd = get_column_letter(price_col_idx_cmd + 1); total_col_letter_cmd = get_column_letter(total_col_idx_cmd + 1); formula_ready_cmd = True
                                     except Exception as e_idx_cmd: logging.error(f"Export CMD: Error get col idx: {e_idx_cmd}")
                                 if formula_ready_cmd:
                                     for supplier_cmd_exp in suppliers_cmd_displayed:
                                         df_supplier_cmd_exp = df_export_cmd[df_export_cmd["Fournisseur"] == supplier_cmd_exp].copy()
                                         if not df_supplier_cmd_exp.empty:
                                             df_sheet_cmd_data = df_supplier_cmd_exp[export_columns_cmd].copy(); num_data_rows_cmd = len(df_sheet_cmd_data)
                                             total_val_cmd = df_sheet_cmd_data[total_col_name_cmd].sum(); req_min_cmd_exp = min_order_dict.get(supplier_cmd_exp, 0); min_fmt_cmd = f"{req_min_cmd_exp:,.2f} €" if req_min_cmd_exp > 0 else "N/A"
                                             if "Désignation Article" in export_columns_cmd: lbl_col_cmd = "Désignation Article";
                                             elif "Référence Article" in export_columns_cmd: lbl_col_cmd = "Référence Article";
                                             else: lbl_col_cmd = export_columns_cmd[1];
                                             total_row_dict_cmd = {c: "" for c in export_columns_cmd}; total_row_dict_cmd[lbl_col_cmd] = "TOTAL COMMANDE"; total_row_dict_cmd[total_col_name_cmd] = total_val_cmd
                                             min_row_dict_cmd = {c: "" for c in export_columns_cmd}; min_row_dict_cmd[lbl_col_cmd] = "Minimum Requis"; min_row_dict_cmd[total_col_name_cmd] = min_fmt_cmd
                                             df_sheet_cmd = pd.concat([df_sheet_cmd_data, pd.DataFrame([total_row_dict_cmd]), pd.DataFrame([min_row_dict_cmd])], ignore_index=True)
                                             sanitized_name_cmd = sanitize_sheet_name(supplier_cmd_exp)
                                             try:
                                                 df_sheet_cmd.to_excel(writer_cmd, sheet_name=sanitized_name_cmd, index=False)
                                                 ws_cmd = writer_cmd.sheets[sanitized_name_cmd]
                                                 for r_num in range(2, num_data_rows_cmd + 2):
                                                     formula = f"={qty_col_letter_cmd}{r_num}*{price_col_letter_cmd}{r_num}"; cell = ws_cmd[f"{total_col_letter_cmd}{r_num}"]; cell.value = formula; cell.number_format = '#,##0.00 €'
                                                 total_formula_row_cmd = num_data_rows_cmd + 2
                                                 if num_data_rows_cmd > 0:
                                                     sum_formula = f"=SUM({total_col_letter_cmd}2:{total_col_letter_cmd}{num_data_rows_cmd + 1})"; sum_cell = ws_cmd[f"{total_col_letter_cmd}{total_formula_row_cmd}"]; sum_cell.value = sum_formula; sum_cell.number_format = '#,##0.00 €'
                                                 sheets_created_cmd += 1
                                             except Exception as write_err_cmd: logging.exception(f"Export CMD: Error write sheet {sanitized_name_cmd}: {write_err_cmd}")
                                 else: st.error("Export CMD: Erreur identification colonnes formules.")
                         except Exception as e_writer_cmd: logging.exception(f"Export CMD: ExcelWriter error: {e_writer_cmd}")

                         if sheets_created_cmd > 0:
                              output_cmd.seek(0)
                              fname_cmd = f"commande_{'multiples' if len(suppliers_cmd_displayed)>1 else sanitize_sheet_name(suppliers_cmd_displayed[0])}_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx"
                              st.download_button(label=f"📥 Télécharger Commande ({sheets_created_cmd} Onglet{'s' if sheets_created_cmd>1 else ''})", data=output_cmd, file_name=fname_cmd, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_cmd_btn")
                         else: st.info("Aucune quantité > 0 à exporter pour la commande calculée.")
                    else: st.info("Aucune quantité > 0 trouvée dans les résultats à exporter.")
                else:
                    st.info("Les résultats affichés précédemment ne correspondent pas à la sélection actuelle de fournisseurs. Veuillez relancer le calcul si nécessaire.")


    # ====================== TAB 2: Analyse Rotation Stock ======================
    with tab2:
        st.header("Analyse de la Rotation des Stocks")
        st.caption("Utilise les fournisseurs sélectionnés dans la barre latérale.")

        data_available_for_analysis = (not df_display_filtered.empty and semaine_columns)

        if not selected_fournisseurs: st.info("Veuillez sélectionner au moins un fournisseur dans la barre latérale.")
        elif df_display_filtered.empty and selected_fournisseurs: st.warning("Aucun article trouvé pour le(s) fournisseur(s) sélectionné(s).")
        elif not semaine_columns: st.warning("Analyse impossible: Aucune colonne de ventes valide identifiée.")
        else:
            st.markdown("#### Paramètres d'Analyse")
            col1_rot, col2_rot = st.columns(2)
            with col1_rot:
                period_options = {"12 dernières semaines": 12, "52 dernières semaines": 52, "Tout l'historique": 0 }
                selected_period_label = st.selectbox("📅 Période calcul ventes:", options=period_options.keys(), key="rotation_period_select")
                selected_period_weeks = period_options[selected_period_label]
            with col2_rot:
                 st.markdown("##### Options d'Affichage")
                 show_all_products = st.checkbox("Afficher tous les produits", value=st.session_state.get('show_all_rotation', True), key="show_all_rotation_cb", help="Si coché, ignore le filtre sur les ventes mensuelles.")
                 st.session_state.show_all_rotation = show_all_products
                 rotation_threshold = st.number_input("... ou afficher ventes mensuelles <", min_value=0.0, value=st.session_state.rotation_threshold_value, step=0.1, format="%.1f", key="rotation_threshold_input", disabled=show_all_products, help="Seuil de vente moyenne mensuelle.")
                 if not show_all_products: st.session_state.rotation_threshold_value = rotation_threshold

            if st.button("🔄 Analyser la Rotation", key="analyze_rotation_button"):
                 with st.spinner("Analyse en cours..."): df_rotation_result = calculer_rotation_stock(df_display_filtered, semaine_columns, selected_period_weeks)
                 if df_rotation_result is not None:
                     st.success("✅ Analyse de rotation terminée."); st.session_state.rotation_result_df = df_rotation_result; st.session_state.rotation_period_label = selected_period_label; st.session_state.selected_fournisseurs_calc_rot = selected_fournisseurs
                     st.rerun()
                 else:
                      st.error("❌ L'analyse de rotation a échoué.")
                      if 'rotation_result_df' in st.session_state: del st.session_state.rotation_result_df

            # Display Rotation Results (Apply filter conditionally)
            if 'rotation_result_df' in st.session_state and st.session_state.rotation_result_df is not None:
                 if st.session_state.selected_fournisseurs_calc_rot == selected_fournisseurs:
                    st.markdown("---"); st.markdown(f"#### Résultats de l'Analyse de Rotation ({st.session_state.get('rotation_period_label', '')})")
                    df_results_rot_orig = st.session_state.rotation_result_df
                    threshold_display = st.session_state.rotation_threshold_value
                    show_all_flag = st.session_state.show_all_rotation

                    # Conditional Filtering
                    monthly_sales_col = "Ventes Moy Mensuel (Période)"; can_filter = False; df_results_rot_to_display = pd.DataFrame()
                    if monthly_sales_col in df_results_rot_orig.columns:
                        monthly_sales_series = pd.to_numeric(df_results_rot_orig[monthly_sales_col], errors='coerce').fillna(0); can_filter = True
                    else: st.warning(f"Colonne '{monthly_sales_col}' non trouvée.")

                    if show_all_flag: df_results_rot_to_display = df_results_rot_orig.copy(); st.caption(f"Affichage de tous les {len(df_results_rot_to_display)} articles.")
                    elif can_filter:
                        try: df_results_rot_to_display = df_results_rot_orig[monthly_sales_series < threshold_display].copy(); st.caption(f"Filtre appliqué : Ventes < {threshold_display:.1f}/mois. {len(df_results_rot_to_display)} / {len(df_results_rot_orig)} articles.")
                        except Exception as e_filter_rot: st.error(f"Erreur filtre : {e_filter_rot}"); df_results_rot_to_display = df_results_rot_orig.copy()
                    else: df_results_rot_to_display = df_results_rot_orig.copy()

                    # Display the Filtered or Unfiltered DataFrame
                    rotation_display_cols = ["AF_RefFourniss", "Référence Article", "Désignation Article", "Tarif d'achat", "Stock", "Unités Vendues (Période)", "Ventes Moy Hebdo (Période)", "Ventes Moy Mensuel (Période)", "Semaines Stock (WoS)", "Rotation Unités (Proxy)", "Valeur Stock Actuel (€)", "COGS (Période)", "Rotation Valeur (Proxy)"]
                    rotation_display_cols_final = [col for col in rotation_display_cols if col in df_results_rot_to_display.columns]

                    if df_results_rot_to_display.empty:
                        if not df_results_rot_orig.empty and not show_all_flag and can_filter: st.info(f"Aucun article < {threshold_display:.1f} ventes/mois.")
                    elif not rotation_display_cols_final: st.error("Aucune colonne rotation trouvée après filtrage.")
                    else:
                        df_rot_display_copy = df_results_rot_to_display[rotation_display_cols_final].copy()
                        numeric_cols_to_round = {"Tarif d'achat": 2, "Ventes Moy Hebdo (Période)": 2, "Ventes Moy Mensuel (Période)": 2, "Semaines Stock (WoS)": 1, "Rotation Unités (Proxy)": 2, "Valeur Stock Actuel (€)": 2, "COGS (Période)": 2, "Rotation Valeur (Proxy)": 2}
                        for col, decimals in numeric_cols_to_round.items():
                            if col in df_rot_display_copy.columns:
                                 df_rot_display_copy[col] = pd.to_numeric(df_rot_display_copy[col], errors='coerce')
                                 if pd.api.types.is_numeric_dtype(df_rot_display_copy[col]): df_rot_display_copy[col] = df_rot_display_copy[col].round(decimals)
                        df_rot_display_copy.replace([np.inf, -np.inf], 'Inf', inplace=True)
                        formatters = {"Tarif d'achat": "{:,.2f}€", "Stock": "{:,.0f}", "Unités Vendues (Période)": "{:,.0f}", "Ventes Moy Hebdo (Période)": "{:,.2f}", "Ventes Moy Mensuel (Période)": "{:,.2f}", "Semaines Stock (WoS)": "{}", "Rotation Unités (Proxy)": "{}", "Valeur Stock Actuel (€)": "{:,.2f}€", "COGS (Période)": "{:,.2f}€", "Rotation Valeur (Proxy)": "{}"}
                        st.dataframe(df_rot_display_copy.style.format(formatters, na_rep="-", thousands=","))

                    # Export Rotation Data (Exports the DISPLAYED data)
                    st.markdown("#### Exportation de l'Analyse Affichée")
                    if not df_results_rot_to_display.empty: # Export based on DISPLAYED results
                         output_rot = io.BytesIO()
                         export_rot_cols_base = ["AF_RefFourniss", "Référence Article", "Désignation Article", "Tarif d'achat", "Stock", "Unités Vendues (Période)", "Ventes Moy Hebdo (Période)", "Ventes Moy Mensuel (Période)", "Semaines Stock (WoS)", "Rotation Unités (Proxy)", "Valeur Stock Actuel (€)", "COGS (Période)", "Rotation Valeur (Proxy)"]
                         export_rot_cols_with_fourn = ["Fournisseur"] + export_rot_cols_base if "Fournisseur" in df_results_rot_to_display.columns else export_rot_cols_base
                         export_rot_cols_final = [col for col in export_rot_cols_with_fourn if col in df_results_rot_to_display.columns]
                         df_export_rot = df_results_rot_to_display[export_rot_cols_final].copy() # Use displayed DF
                         for col, decimals in numeric_cols_to_round.items(): # Reuse rounding dict
                              if col in df_export_rot.columns:
                                  df_export_rot[col] = pd.to_numeric(df_export_rot[col], errors='coerce')
                                  if pd.api.types.is_numeric_dtype(df_export_rot[col]): df_export_rot[col] = df_export_rot[col].round(decimals)
                         df_export_rot.replace([np.inf, -np.inf], 'Infini', inplace=True) # Replace inf AFTER rounding
                         export_label = f"Filtree_{threshold_display:.1f}" if not show_all_flag else "Complete"
                         sheet_name_rot = f"Rotation_{export_label}"
                         fname_rot_base = f"analyse_rotation_{export_label}"
                         with pd.ExcelWriter(output_rot, engine="openpyxl") as writer_rot: df_export_rot.to_excel(writer_rot, sheet_name=sheet_name_rot, index=False)
                         output_rot.seek(0)
                         suppliers_export_rot = st.session_state.get('selected_fournisseurs_session', [])
                         fname_rot = f"{fname_rot_base}_{'multiples' if len(suppliers_export_rot)>1 else sanitize_sheet_name(suppliers_export_rot[0] if suppliers_export_rot else 'NA')}_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx"
                         download_label_rot = f"📥 Télécharger Analyse {'Filtrée' if not show_all_flag else 'Complète'}" + (f" (<{threshold_display:.1f}/mois)" if not show_all_flag else "")
                         st.download_button(label=download_label_rot, data=output_rot, file_name=fname_rot, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_rot_btn")
                    elif not df_results_rot_orig.empty: st.info(f"Aucune donnée de rotation correspondant aux critères actuels à exporter.")
                    else: st.info("Aucune donnée de rotation calculée à exporter.")
                 else: # Results in session state don't match current selection
                     st.info("Les résultats d'analyse affichés précédemment ne correspondent pas à la sélection actuelle de fournisseurs. Veuillez relancer l'analyse si nécessaire.")


    # ========================= TAB 3: Vérification Stock =========================
    with tab3:
        st.header("Vérification des Stocks Négatifs")
        st.caption("Analyse tous les articles du fichier chargé ('Tableau final').")

        df_source_for_neg_stock = st.session_state.get('df_full', None)

        if df_source_for_neg_stock is None: st.warning("Données non chargées.")
        elif df_source_for_neg_stock.empty: st.warning("Aucune donnée dans 'Tableau final'.")
        else:
            stock_col = "Stock"
            if stock_col not in df_source_for_neg_stock.columns: st.error(f"Colonne '{stock_col}' non trouvée.")
            else:
                # Stock column already numeric from load
                df_stock_negatif = df_source_for_neg_stock[df_source_for_neg_stock[stock_col] < 0].copy()
                if df_stock_negatif.empty: st.success("✅ Aucune anomalie de stock négatif détectée.")
                else:
                    st.warning(f"⚠️ **{len(df_stock_negatif)} article(s) avec stock négatif détecté(s) !**")
                    neg_stock_display_cols = ["Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article", "Stock"]
                    neg_stock_display_cols_final = [col for col in neg_stock_display_cols if col in df_stock_negatif.columns]
                    if not neg_stock_display_cols_final: st.error("Colonnes manquantes affichage stocks négatifs.")
                    else: st.dataframe(df_stock_negatif[neg_stock_display_cols_final].style.format({"Stock": "{:,.0f}"}, na_rep="-").apply(lambda x: ['background-color: #FADBD8' if v < 0 else '' for v in x], subset=['Stock']))
                    st.markdown("---"); st.markdown("#### Exporter la Liste Complète des Stocks Négatifs")
                    output_neg = io.BytesIO(); df_export_neg = df_stock_negatif[neg_stock_display_cols_final].copy()
                    try:
                        with pd.ExcelWriter(output_neg, engine="openpyxl") as writer_neg: df_export_neg.to_excel(writer_neg, sheet_name="Stocks_Negatifs_Complets", index=False)
                        output_neg.seek(0); fname_neg = f"stocks_negatifs_complets_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx"
                        st.download_button(label="📥 Télécharger Liste Stocks Négatifs (Tous)", data=output_neg, file_name=fname_neg, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_neg_stock_btn")
                    except Exception as e_export_neg: st.error(f"Erreur export stocks négatifs: {e_export_neg}"); logging.exception("Error exporting negative stocks:")

    # ========================= TAB 4: Simulation Forecast =========================
    with tab4:
        st.header("Simulation Forecast Annuel")
        st.caption("Utilise les fournisseurs sélectionnés dans la barre latérale et suppose que les 52 dernières colonnes de ventes représentent l'année N-1.")
        st.warning("🚨 **Approximation Importante:** Le calcul de saisonnalité mensuelle est basé sur un découpage approximatif des 52 dernières colonnes hebdomadaires. Pour une précision accrue, un fichier avec des dates explicites serait nécessaire.")

        # Check conditions for forecast simulation
        if df_display_filtered.empty:
             if selected_fournisseurs: st.warning("Aucun article trouvé pour le(s) fournisseur(s) sélectionné(s).")
             else: st.info("Veuillez sélectionner au moins un fournisseur.")
        elif len(semaine_columns) < 52:
            st.warning("Données historiques insuffisantes (< 52 semaines) pour cette simulation.")
        else:
            st.markdown("#### Paramètres de Simulation")

            # Select Months
            all_months = list(calendar.month_name)[1:]
            # Use session state to remember selected months or default to all
            default_months = st.session_state.get('forecast_selected_months', all_months)
            selected_months_forecast = st.multiselect(
                "📅 Mois à inclure dans la simulation:",
                options=all_months, default=default_months, key="forecast_months_select"
            )
            st.session_state.forecast_selected_months = selected_months_forecast # Store selection

            # Simulation Type
            sim_type = st.radio(
                "⚙️ Type de Simulation:", ('Simple Progression', 'Objectif Montant'),
                key="forecast_sim_type", horizontal=True,
                index=0 if st.session_state.get('forecast_sim_type_index', 0) == 0 else 1 # Persist radio choice
            )
            st.session_state.forecast_sim_type_index = 0 if sim_type == 'Simple Progression' else 1

            # Conditional Inputs
            progression_pct = 0.0; objectif_montant = 0.0
            col1_fcst, col2_fcst = st.columns(2)
            with col1_fcst:
                if sim_type == 'Simple Progression':
                    progression_pct = st.number_input( "📈 Progression vs N-1 (%)", -100.0, value=st.session_state.get('forecast_prog_pct', 5.0), step=0.5, format="%.1f", key="forecast_prog_pct")
                    st.session_state.forecast_prog_pct = progression_pct
            with col2_fcst:
                if sim_type == 'Objectif Montant':
                    objectif_montant = st.number_input( "🎯 Objectif Montant Total (€)", min_value=0.0, value=st.session_state.get('forecast_target_amount', 10000.0), step=1000.0, format="%.2f", key="forecast_target_amount")
                    st.session_state.forecast_target_amount = objectif_montant

            # Simulation Button
            if st.button("▶️ Lancer la Simulation Forecast", key="run_forecast_sim"):
                 if not selected_months_forecast: st.warning("Veuillez sélectionner au moins un mois.")
                 else:
                    with st.spinner("Simulation en cours..."):
                         df_forecast_result = calculer_forecast_simulation(df_display_filtered, semaine_columns, selected_months_forecast, sim_type, progression_pct, objectif_montant)
                    if df_forecast_result is not None:
                        st.success("✅ Simulation terminée."); st.session_state.forecast_result_df = df_forecast_result
                        st.session_state.forecast_params = {'suppliers': selected_fournisseurs, 'months': selected_months_forecast, 'type': sim_type, 'prog': progression_pct, 'obj': objectif_montant}
                        st.rerun()
                    else:
                        st.error("❌ La simulation Forecast a échoué.")
                        if 'forecast_result_df' in st.session_state: del st.session_state.forecast_result_df

            # Display Forecast Results
            if 'forecast_result_df' in st.session_state and st.session_state.forecast_result_df is not None:
                 current_params = {'suppliers': selected_fournisseurs, 'months': selected_months_forecast, 'type': sim_type, 'prog': progression_pct, 'obj': objectif_montant}
                 if st.session_state.get('forecast_params') == current_params:
                    st.markdown("---"); st.markdown("#### Résultats de la Simulation Forecast")
                    df_results_fcst_display = st.session_state.forecast_result_df

                    month_qty_cols = [f"Qté Prév. {m}" for m in selected_months_forecast if f"Qté Prév. {m}" in df_results_fcst_display.columns]
                    month_amt_cols = [f"Montant Prév. {m} (€)" for m in selected_months_forecast if f"Montant Prév. {m} (€)" in df_results_fcst_display.columns]
                    n1_month_cols_display = [f"Ventes N-1 {m}" for m in selected_months_forecast if f"Ventes N-1 {m}" in df_results_fcst_display.columns]

                    fcst_id_cols = ["Fournisseur", "Référence Article", "Désignation Article", "Conditionnement", "Tarif d'achat"]; fcst_total_cols = ["Ventes Totales N-1", "Qté Totale Prév.", "Montant Total Prév. (€)"]
                    fcst_display_cols = fcst_id_cols + fcst_total_cols + n1_month_cols_display + month_qty_cols + amt_cols
                    fcst_display_cols_final = [col for col in fcst_display_cols if col in df_results_fcst_display.columns]

                    if df_results_fcst_display.empty: st.info("Aucun résultat à afficher.")
                    elif not fcst_display_cols_final: st.error("Erreur: Colonnes de résultats non trouvées.")
                    else:
                        fcst_formatters = {"Tarif d'achat": "{:,.2f}€", "Conditionnement": "{:,.0f}", "Ventes Totales N-1": "{:,.0f}", "Qté Totale Prév.": "{:,.0f}", "Montant Total Prév. (€)": "{:,.2f}€"}
                        for col in n1_month_cols_display: fcst_formatters[col] = "{:,.0f}"
                        for col in month_qty_cols: fcst_formatters[col] = "{:,.0f}"
                        for col in month_amt_cols: fcst_formatters[col] = "{:,.2f}€"
                        st.dataframe(df_results_fcst_display[fcst_display_cols_final].style.format(fcst_formatters, na_rep="-", thousands=","))

                        # Export Forecast Results
                        st.markdown("#### Exportation de la Simulation Forecast")
                        output_fcst = io.BytesIO(); df_export_fcst = df_results_fcst_display[fcst_display_cols_final].copy()
                        try:
                            with pd.ExcelWriter(output_fcst, engine="openpyxl") as writer_fcst: df_export_fcst.to_excel(writer_fcst, sheet_name=f"Forecast_{sim_type.replace(' ','_')}", index=False)
                            output_fcst.seek(0)
                            fname_fcst_base = f"forecast_{sim_type.replace(' ','_').lower()}"
                            suppliers_fcst_str = st.session_state.get('selected_fournisseurs_session', [])
                            fname_fcst = f"{fname_fcst_base}_{'multiples' if len(suppliers_fcst_str)>1 else sanitize_sheet_name(suppliers_fcst_str[0] if suppliers_fcst_str else 'NA')}_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx"
                            st.download_button(label="📥 Télécharger Simulation Forecast", data=output_fcst, file_name=fname_fcst, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_fcst_btn")
                        except Exception as e_export_fcst: st.error(f"Erreur export forecast: {e_export_fcst}")
                 else: st.info("Les résultats de simulation affichés précédemment ne correspondent pas aux paramètres actuels. Veuillez relancer la simulation.")


# --- App footer/initial message ---
elif not uploaded_file:
    st.info("👋 Bienvenue ! Chargez votre fichier Excel pour commencer.")
    if st.button("🔄 Réinitialiser l'application"):
         keys_to_clear = list(st.session_state.keys())
         for key in keys_to_clear: del st.session_state[key]
         st.rerun()
