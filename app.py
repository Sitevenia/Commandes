# --- START OF FINAL COMPLETE CORRECTED app.py (v13 - Integrated Special Tab, Debug Additions) ---

import streamlit as st
import pandas as pd
import numpy as np
import io
import logging
import re
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment
import calendar
import zipfile
from datetime import timedelta

# --- AI Model Import ---
try:
    from prophet import Prophet
    PROPHET_AVAILABLE = True
except ImportError:
    PROPHET_AVAILABLE = False
    logging.warning("Prophet library not found. AI forecasting will be disabled.")

# --- Logging Configuration ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- Helper to suppress Prophet's verbose output ---
import os
import sys
class SuppressStdoutStderr:
    def __enter__(self):
        self.old_stdout = sys.stdout
        self.old_stderr = sys.stderr
        sys.stdout = open(os.devnull, 'w')
        sys.stderr = open(os.devnull, 'w')
    def __exit__(self, exc_type, exc_val, exc_tb):
        if hasattr(sys.stdout, 'close') and not getattr(sys.stdout, 'closed', True):
             try: sys.stdout.close()
             except Exception: pass
        sys.stdout = self.old_stdout
        if hasattr(sys.stderr, 'close') and not getattr(sys.stderr, 'closed', True):
            try: sys.stderr.close()
            except Exception: pass
        sys.stderr = self.old_stderr

# --- ALL HELPER FUNCTIONS DEFINITIONS ---

def sanitize_sheet_name(name):
    if not isinstance(name, str): name = str(name)
    s_name = re.sub(r'[\[\]:*?/\\<>|"]', '_', name)
    if s_name.startswith("'"): s_name = "_" + s_name[1:]
    if s_name.endswith("'"): s_name = s_name[:-1] + "_"
    return s_name[:31]

def sanitize_supplier_key(supplier_name_str):
    if not isinstance(supplier_name_str, str): supplier_name_str = str(supplier_name_str)
    s_key = re.sub(r'\W+', '_', supplier_name_str)
    s_key = re.sub(r'^_+|_+$', '', s_key)
    s_key = re.sub(r'_+', '_', s_key)
    return s_key if s_key else "invalid_supplier_key_name"

def render_supplier_checkboxes(tab_key_prefix, all_suppliers_list, default_select_all=False):
    select_all_key = f"{tab_key_prefix}_select_all_suppliers"
    supplier_checkbox_keys = { sup: f"{tab_key_prefix}_supplier_cb_{sanitize_supplier_key(sup)}" for sup in all_suppliers_list }

    if select_all_key not in st.session_state: st.session_state[select_all_key] = default_select_all
    for cb_key in supplier_checkbox_keys.values():
        if cb_key not in st.session_state: st.session_state[cb_key] = st.session_state[select_all_key]

    def toggle_all_suppliers_for_tab():
        current_val = st.session_state[select_all_key]
        for cb_k_val in supplier_checkbox_keys.values(): st.session_state[cb_k_val] = current_val

    def check_individual_supplier_for_tab():
        all_checked = all(st.session_state.get(cb_k_val, False) for cb_k_val in supplier_checkbox_keys.values())
        if st.session_state.get(select_all_key) != all_checked:
            st.session_state[select_all_key] = all_checked

    exp_label = "👤 Sélectionner Fournisseurs"
    if tab_key_prefix == "tab5": exp_label = "👤 Sélectionner Fournisseurs pour Export Suivi Commandes"

    with st.expander(exp_label, expanded=True):
        st.checkbox("Sélectionner / Désélectionner Tout", key=select_all_key, on_change=toggle_all_suppliers_for_tab, disabled=not bool(all_suppliers_list))
        st.markdown("---")
        selected_suppliers_ui = []
        num_cols = 4; checkbox_cols = st.columns(num_cols); col_idx = 0
        for sup_name, cb_k_val in supplier_checkbox_keys.items():
            with checkbox_cols[col_idx]:
                st.checkbox(sup_name, key=cb_k_val, on_change=check_individual_supplier_for_tab)
            if st.session_state.get(cb_k_val): selected_suppliers_ui.append(sup_name)
            col_idx = (col_idx + 1) % num_cols
    return selected_suppliers_ui

def safe_read_excel(uploaded_file, sheet_name, **kwargs):
    try:
        if isinstance(uploaded_file, io.BytesIO): uploaded_file.seek(0)
        file_name_attr = getattr(uploaded_file, 'name', '')
        engine_to_use = 'openpyxl' if file_name_attr.lower().endswith('.xlsx') else None
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, engine=engine_to_use, **kwargs)
        if df is None: return None
        if df.empty and len(df.columns) == 0 and sheet_name is not None: return pd.DataFrame()
        return df
    except ValueError as e:
        if f"Worksheet named '{sheet_name}' not found" in str(e) or f"'{sheet_name}' not found" in str(e):
             st.warning(f"⚠️ Onglet '{sheet_name}' non trouvé.") 
        else: st.error(f"❌ Erreur valeur lecture onglet '{sheet_name}': {e}.")
        return None
    except Exception as e:
        if "zip file" in str(e).lower() or "BadZipFile" in str(type(e).__name__):
             st.error(f"❌ Erreur lecture onglet '{sheet_name}': Fichier .xlsx corrompu.")
        else: st.error(f"❌ Erreur inattendue ('{type(e).__name__}') lecture onglet '{sheet_name}': {e}.")
        return None

def format_excel_sheet(worksheet, df, column_formats={}, freeze_header=True, default_float_format="#,##0.00", default_int_format="#,##0", default_date_format="dd/mm/yyyy"):
    if df is None or df.empty: return
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    data_alignment = Alignment(vertical="center")
    for cell in worksheet[1]: cell.font, cell.fill, cell.alignment = header_font, header_fill, header_alignment
    for idx, col_name in enumerate(df.columns):
        col_letter = get_column_letter(idx + 1); num_fmt_apply = None
        try:
            hdr_len = len(str(col_name))
            non_na_s = df[col_name].dropna()
            samp_data = non_na_s.sample(min(len(non_na_s), 20)) if not non_na_s.empty else pd.Series([], dtype='object')
            data_len = samp_data.astype(str).map(len).max() if not samp_data.empty else 0
            data_len = data_len if pd.notna(data_len) else 0
            max_len = min(max(max(hdr_len, data_len) + 3, 10), 50)
            worksheet.column_dimensions[col_letter].width = max_len
        except Exception: worksheet.column_dimensions[col_letter].width = 15
        spec_fmt = column_formats.get(col_name)
        try: col_dtype = df[col_name].dtype
        except KeyError: continue
        if spec_fmt: num_fmt_apply = spec_fmt
        elif pd.api.types.is_integer_dtype(col_dtype): num_fmt_apply = default_int_format
        elif pd.api.types.is_float_dtype(col_dtype): num_fmt_apply = default_float_format
        elif pd.api.types.is_datetime64_any_dtype(col_dtype) or (not df[col_name].empty and isinstance(df[col_name].dropna().iloc[0] if not df[col_name].dropna().empty else None, pd.Timestamp)):
            num_fmt_apply = default_date_format
        for row_idx in range(2, worksheet.max_row + 1):
            cell = worksheet[f"{col_letter}{row_idx}"]
            cell.alignment = data_alignment
            if num_fmt_apply and cell.value is not None and not str(cell.value).startswith('='):
                try: cell.number_format = num_fmt_apply
                except Exception: pass
    if freeze_header: worksheet.freeze_panes = worksheet['A2']

def calculer_quantite_a_commander(df, semaine_columns, montant_minimum_input, duree_semaines):
    try:
        if not isinstance(df, pd.DataFrame) or df.empty: st.info("Aucune donnée pour calcul qtés."); return None
        req_cols=["Stock","Conditionnement","Tarif d'achat"]+semaine_columns
        miss_cols=[c for c in req_cols if c not in df.columns]
        if miss_cols: st.error(f"Cols manquantes (calcul): {', '.join(miss_cols)}"); return None
        if not semaine_columns: st.error("Aucune col 'semaine' identifiée (calcul)."); return None
        df_calc=df.copy()
        for col in req_cols: df_calc[col]=pd.to_numeric(df_calc[col],errors='coerce').replace([np.inf,-np.inf],np.nan).fillna(0)
        num_sem_tot=len(semaine_columns); ventes_N1=df_calc[semaine_columns].sum(axis=1)
        if num_sem_tot>=64:
            v12N1=df_calc[semaine_columns[-64:-52]].sum(axis=1); v12N1s=df_calc[semaine_columns[-52:-40]].sum(axis=1)
            avg12N1=v12N1/12; avg12N1s=v12N1s/12
        else: v12N1,v12N1s,avg12N1,avg12N1s=(pd.Series(0.0,index=df_calc.index)for _ in range(4))
        nb_sem_rec=min(num_sem_tot,12)
        if nb_sem_rec>0: v12last=df_calc[semaine_columns[-nb_sem_rec:]].sum(axis=1); avg12last=v12last/nb_sem_rec
        else: v12last,avg12last=(pd.Series(0.0,index=df_calc.index)for _ in range(2))
        qpond=(0.5*avg12last+0.2*avg12N1+0.3*avg12N1s); qnec=qpond*duree_semaines
        qcomm_s=(qnec-df_calc["Stock"]).apply(lambda x:max(0,x))
        cond,stock,tarif=df_calc["Conditionnement"],df_calc["Stock"],df_calc["Tarif d'achat"]
        qcomm=qcomm_s.tolist()
        for i in range(len(qcomm)):
            c,q=cond.iloc[i],qcomm[i]
            if q>0 and c>0: qcomm[i]=int(np.ceil(q/c)*c)
            elif q>0 and c<=0: logging.warning(f"Art idx {df_calc.index[i]} (Ref: {df_calc.get('Référence Article',pd.Series(['N/A'], index=df_calc.index)).iloc[i]}) Qté {q:.2f} ignorée car cond={c}."); qcomm[i]=0
            else: qcomm[i]=0
        if nb_sem_rec>0:
            for i in range(len(qcomm)):
                c=cond.iloc[i]; vr_cnt=(df_calc[semaine_columns[-nb_sem_rec:]].iloc[i]>0).sum()
                if vr_cnt>=2 and stock.iloc[i]<=1 and c>0: qcomm[i]=max(qcomm[i],c)
        for i in range(len(qcomm)):
            vt_n1_it,vr_sum_it=ventes_N1.iloc[i],v12last.iloc[i]
            if vt_n1_it<6 and vr_sum_it<2: qcomm[i]=0
        qcomm_df_t=pd.Series(qcomm,index=df_calc.index); mt_avant_adj= (qcomm_df_t*tarif).sum()
        if montant_minimum_input>0 and mt_avant_adj<montant_minimum_input:
            mt_act=mt_avant_adj; elig_incr=[]
            for i in range(len(qcomm)):
                if qcomm[i]>0 and cond.iloc[i]>0 and tarif.iloc[i]>0: elig_incr.append(i)
            if not elig_incr:
                if mt_act<montant_minimum_input: st.warning(f"Impossible atteindre min {montant_minimum_input:,.2f}€. Actuel: {mt_act:,.2f}€. Aucun article éligible.")
            else:
                idx_ptr_el=0; max_iter_l=len(elig_incr)*20+1; iters=0
                while mt_act<montant_minimum_input and iters<max_iter_l:
                    iters+=1; orig_df_idx_pos=elig_incr[idx_ptr_el]
                    c_it,p_it=cond.iloc[orig_df_idx_pos],tarif.iloc[orig_df_idx_pos]
                    if c_it > 0 and p_it > 0:
                        qcomm[orig_df_idx_pos]+=c_it; mt_act+=c_it*p_it
                    idx_ptr_el=(idx_ptr_el+1)%len(elig_incr)
                if iters>=max_iter_l and mt_act<montant_minimum_input: st.error(f"Ajustement min: Max iter ({max_iter_l}) atteint. Actuel: {mt_act:,.2f}€ / Requis: {montant_minimum_input:,.2f}€.")
        qcomm_fin_s=pd.Series(qcomm,index=df_calc.index); mt_fin=(qcomm_fin_s*tarif).sum()
        return(qcomm_fin_s,ventes_N1,v12N1,v12last,mt_fin)
    except KeyError as e:st.error(f"Err clé (calcul qtés): '{e}'.");logging.exception(f"KeyError in calc_qte_cmd: {e}");return None
    except Exception as e:st.error(f"Err inattendue (calcul qtés): {type(e).__name__} - {e}");logging.exception("Exception in calc_qte_cmd:");return None

def calculer_rotation_stock(df, semaine_columns, periode_semaines_analyse):
    try:
        if not isinstance(df, pd.DataFrame) or df.empty: st.info("Aucune donnée pour analyse rotation."); return pd.DataFrame()
        req_cols=["Stock","Tarif d'achat"]; miss_cols=[c for c in req_cols if c not in df.columns]
        if miss_cols: st.error(f"Cols manquantes (rotation): {', '.join(miss_cols)}"); return None
        df_rot=df.copy()
        if periode_semaines_analyse and periode_semaines_analyse>0 and len(semaine_columns)>=periode_semaines_analyse:sem_an,nb_sem_an=semaine_columns[-periode_semaines_analyse:],periode_semaines_analyse
        elif periode_semaines_analyse and periode_semaines_analyse>0:sem_an,nb_sem_an=semaine_columns,len(semaine_columns);st.caption(f"Période analyse ajustée à {nb_sem_an} sem.")
        else:sem_an,nb_sem_an=semaine_columns,len(semaine_columns)
        if not sem_an:
            st.warning("Aucune col vente pour analyse rotation.")
            metric_cols=["Unités Vendues (Période)","Ventes Moy Hebdo (Période)","Ventes Moy Mensuel (Période)","Semaines Stock (WoS)","Rotation Unités (Proxy)","COGS (Période)","Valeur Stock Actuel (€)","Rotation Valeur (Proxy)"]
            for m_col in metric_cols:df_rot[m_col]=0.0
            return df_rot
        for col in sem_an:
            if col in df_rot.columns: df_rot[col]=pd.to_numeric(df_rot[col],errors='coerce').fillna(0)
            else: df_rot[col] = 0.0
        df_rot["Unités Vendues (Période)"]=df_rot[[c for c in sem_an if c in df_rot.columns]].sum(axis=1)
        df_rot["Ventes Moy Hebdo (Période)"]=df_rot["Unités Vendues (Période)"]/nb_sem_an if nb_sem_an>0 else 0.0
        df_rot["Ventes Moy Mensuel (Période)"]=df_rot["Ventes Moy Hebdo (Période)"]*(52/12.0)
        df_rot["Stock"]=pd.to_numeric(df_rot["Stock"],errors='coerce').fillna(0)
        df_rot["Tarif d'achat"]=pd.to_numeric(df_rot["Tarif d'achat"],errors='coerce').fillna(0)
        den_wos=df_rot["Ventes Moy Hebdo (Période)"]
        df_rot["Semaines Stock (WoS)"]=np.divide(df_rot["Stock"],den_wos,out=np.full_like(df_rot["Stock"],np.inf,dtype=np.float64),where=den_wos!=0)
        df_rot.loc[df_rot["Stock"]<=0,"Semaines Stock (WoS)"]=0.0
        den_rot_u=df_rot["Stock"]
        df_rot["Rotation Unités (Proxy)"]=np.divide(df_rot["Unités Vendues (Période)"],den_rot_u,out=np.full_like(den_rot_u,np.inf,dtype=np.float64),where=den_rot_u!=0)
        df_rot.loc[(df_rot["Unités Vendues (Période)"]<=0)&(den_rot_u<=0),"Rotation Unités (Proxy)"]=0.0
        df_rot.loc[(df_rot["Unités Vendues (Période)"]<=0)&(den_rot_u>0),"Rotation Unités (Proxy)"]=0.0
        df_rot["COGS (Période)"]=df_rot["Unités Vendues (Période)"]*df_rot["Tarif d'achat"]
        df_rot["Valeur Stock Actuel (€)"]=df_rot["Stock"]*df_rot["Tarif d'achat"]
        den_rot_v=df_rot["Valeur Stock Actuel (€)"]
        df_rot["Rotation Valeur (Proxy)"]=np.divide(df_rot["COGS (Période)"],den_rot_v,out=np.full_like(den_rot_v,np.inf,dtype=np.float64),where=den_rot_v!=0)
        df_rot.loc[(df_rot["COGS (Période)"]<=0)&(den_rot_v<=0),"Rotation Valeur (Proxy)"]=0.0
        df_rot.loc[(df_rot["COGS (Période)"]<=0)&(den_rot_v>0),"Rotation Valeur (Proxy)"]=0.0
        return df_rot
    except KeyError as e:st.error(f"Err clé (rotation): '{e}'.");logging.exception(f"KeyError in calc_rotation: {e}");return None
    except Exception as e:st.error(f"Err inattendue (rotation): {type(e).__name__} - {e}");logging.exception("Error in calc_rotation:");return None

def approx_weeks_to_months(week_cols_52_names):
    month_map = {}
    if not week_cols_52_names or len(week_cols_52_names) != 52:
        logging.warning(f"approx_weeks_to_months expects 52 week col names, got {len(week_cols_52_names) if week_cols_52_names else 0}.")
        return month_map
    weeks_per_month_approx = 52.0 / 12.0
    for i in range(1, 13):
        month_name = calendar.month_name[i]
        start_week_index = int(round((i - 1) * weeks_per_month_approx))
        end_week_index = int(round(i * weeks_per_month_approx))
        start_week_index = max(0, start_week_index)
        end_week_index = min(52, end_week_index)
        if start_week_index < end_week_index:
            month_map[month_name] = week_cols_52_names[start_week_index:end_week_index]
        else: month_map[month_name] = []
    logging.info(f"Approx month map created. Ex: January: {month_map.get('January', [])}")
    return month_map

def calculer_forecast_simulation_v3(df, all_historical_semaine_columns, selected_months_list, sim_type_str, progression_pct_val=0, objectif_montant_val=0):
    try:
        if not isinstance(df, pd.DataFrame) or df.empty: st.warning("Aucune donnée pour simu forecast."); return None,0.0
        if not all_historical_semaine_columns or len(all_historical_semaine_columns)<52: st.error("Données histo. < 52 sem. pour N-1."); return None,0.0
        if not selected_months_list: st.warning("Sélectionner au moins un mois pour simu."); return None,0.0
        required_data_cols = ["Référence Article","Désignation Article","Conditionnement","Tarif d'achat","Fournisseur"]
        if not all(c in df.columns for c in required_data_cols): st.error(f"Cols manquantes (simu): {', '.join([c for c in required_data_cols if c not in df.columns])}"); return None,0.0
        parsed_week_col_objects = []
        available_years = set()
        for col_name_str in all_historical_semaine_columns:
            if isinstance(col_name_str, str):
                match = re.match(r"(\d{4})[SW]?(\d{1,2})", col_name_str, re.IGNORECASE)
                if match:
                    year, week_num = int(match.group(1)), int(match.group(2))
                    if 1 <= week_num <= 53: available_years.add(year); parsed_week_col_objects.append({'year': year, 'week': week_num, 'col': col_name_str, 'sort_key': year * 100 + week_num})
        if not available_years: st.error("Impossible de déterminer années. Format: 'YYYYWW' ou 'YYYYSwW'."); return None,0.0
        parsed_week_col_objects.sort(key=lambda x: x['sort_key'])
        current_year_n = max(available_years) if available_years else 0
        previous_year_n_minus_1 = current_year_n - 1
        st.caption(f"Simu N-1 (N: {current_year_n}, N-1: {previous_year_n_minus_1})")
        n1_week_data_objects = [item for item in parsed_week_col_objects if item['year'] == previous_year_n_minus_1]
        if len(n1_week_data_objects) < 52: st.error(f"Données N-1 ({previous_year_n_minus_1}) < 52 sem. ({len(n1_week_data_objects)})."); return None,0.0
        n1_week_column_names_for_mapping = [item['col'] for item in n1_week_data_objects[:52]]
        df_simulation_results = df[required_data_cols].copy()
        df_simulation_results["Tarif d'achat"] = pd.to_numeric(df_simulation_results["Tarif d'achat"], errors='coerce').fillna(0)
        df_simulation_results["Conditionnement"] = pd.to_numeric(df_simulation_results["Conditionnement"], errors='coerce').fillna(1).apply(lambda x: 1 if x <= 0 else int(x))
        if not all(c in df.columns for c in n1_week_column_names_for_mapping): st.error(f"Err interne: Cols N-1 mappées non trouvées."); return None,0.0
        df_n1_sales_only = df[n1_week_column_names_for_mapping].copy()
        for col_n1 in n1_week_column_names_for_mapping:
            if col_n1 in df_n1_sales_only.columns: df_n1_sales_only[col_n1] = pd.to_numeric(df_n1_sales_only[col_n1], errors='coerce').fillna(0)
            else: df_n1_sales_only[col_n1] = 0.0
        month_to_n1_week_cols_map = approx_weeks_to_months(n1_week_column_names_for_mapping)
        total_n1_sales_for_selected_months_series = pd.Series(0.0, index=df_simulation_results.index)
        monthly_n1_sales_map_for_selected_months = {}
        for month_name_iter in selected_months_list:
            sales_n1_this_month = pd.Series(0.0, index=df_simulation_results.index)
            if month_name_iter in month_to_n1_week_cols_map and month_to_n1_week_cols_map[month_name_iter]:
                actual_cols_for_month_sum = [c for c in month_to_n1_week_cols_map[month_name_iter] if c in df_n1_sales_only.columns]
                if actual_cols_for_month_sum: sales_n1_this_month = df_n1_sales_only[actual_cols_for_month_sum].sum(axis=1)
            monthly_n1_sales_map_for_selected_months[month_name_iter] = sales_n1_this_month
            total_n1_sales_for_selected_months_series += sales_n1_this_month
            df_simulation_results[f"Ventes N-1 {month_name_iter}"] = sales_n1_this_month
        df_simulation_results["Vts N-1 Tot (Mois Sel.)"] = total_n1_sales_for_selected_months_series
        period_seasonality_factors_map = {}
        safe_total_n1_sales_for_factors = total_n1_sales_for_selected_months_series.copy()
        for month_name_iter in selected_months_list:
            n1_sales_for_month = monthly_n1_sales_map_for_selected_months.get(month_name_iter, pd.Series(0.0, index=df_simulation_results.index))
            factor = np.divide(n1_sales_for_month, safe_total_n1_sales_for_factors, out=np.zeros_like(n1_sales_for_month, dtype=float), where=safe_total_n1_sales_for_factors != 0)
            period_seasonality_factors_map[month_name_iter] = pd.Series(factor, index=df_simulation_results.index).fillna(0)
        base_monthly_forecast_qty_map = {}
        if sim_type_str == 'Simple Progression':
            progression_factor = 1 + (progression_pct_val / 100.0)
            total_forecasted_qty_for_period = total_n1_sales_for_selected_months_series * progression_factor
            for m_name_fcst in selected_months_list:
                seasonality_factor_series = period_seasonality_factors_map.get(m_name_fcst, pd.Series(0.0, index=df_simulation_results.index))
                base_monthly_forecast_qty_map[m_name_fcst] = total_forecasted_qty_for_period * seasonality_factor_series
        elif sim_type_str == 'Objectif Montant':
            if objectif_montant_val <= 0: st.error("Objectif Montant > 0 requis."); return None,0.0
            total_n1_value_all_selected_months = (total_n1_sales_for_selected_months_series * df_simulation_results["Tarif d'achat"]).sum()
            if total_n1_value_all_selected_months <= 0:
                st.warning("Ventes N-1 (valeur) nulles. Répartition égale du montant objectif.")
                num_selected_m = len(selected_months_list)
                if num_selected_m == 0: return None, 0.0
                num_items_gt_zero_price = (df_simulation_results["Tarif d'achat"] > 0).sum()
                target_amount_per_month_item_avg = 0.0
                if num_items_gt_zero_price > 0: target_amount_per_month_item_avg = objectif_montant_val / num_selected_m / num_items_gt_zero_price
                else: st.warning("Aucun article avec prix > 0. Impossible de répartir objectif montant.")
                for m_name_fcst in selected_months_list:
                    base_monthly_forecast_qty_map[m_name_fcst] = np.divide(target_amount_per_month_item_avg,df_simulation_results["Tarif d'achat"],out=np.zeros_like(df_simulation_results["Tarif d'achat"],dtype=float),where=df_simulation_results["Tarif d'achat"]!=0)
            else:
                for m_name_fcst in selected_months_list:
                    monthly_n1_value_series = (monthly_n1_sales_map_for_selected_months.get(m_name_fcst, pd.Series(0.0, index=df_simulation_results.index)) * df_simulation_results["Tarif d'achat"])
                    month_value_contribution_factor = np.divide(monthly_n1_value_series.sum(), total_n1_value_all_selected_months, out=np.array([0.0]), where=total_n1_value_all_selected_months !=0)[0]
                    target_amount_this_month_global = objectif_montant_val * month_value_contribution_factor
                    item_contribution_in_month_value_factor = np.divide(monthly_n1_value_series, monthly_n1_value_series.sum(), out=np.zeros_like(monthly_n1_value_series,dtype=float), where=monthly_n1_value_series.sum() !=0)
                    target_amount_per_item_this_month = target_amount_this_month_global * item_contribution_in_month_value_factor
                    base_monthly_forecast_qty_map[m_name_fcst] = np.divide(target_amount_per_item_this_month,df_simulation_results["Tarif d'achat"],out=np.zeros_like(df_simulation_results["Tarif d'achat"],dtype=float),where=df_simulation_results["Tarif d'achat"]!=0)
        else:st.error(f"Type simu non reconnu: '{sim_type_str}'.");return None,0.0
        total_adjusted_qty_all_months_series = pd.Series(0.0,index=df_simulation_results.index)
        total_final_amount_all_months_series = pd.Series(0.0,index=df_simulation_results.index)
        for m_name_fcst in selected_months_list:
            forecast_qty_col_name,fcst_amt_col=f"Qté Prév. {m_name_fcst}",f"Montant Prév. {m_name_fcst} (€)"
            base_qty_series=base_monthly_forecast_qty_map.get(m_name_fcst,pd.Series(0.0,index=df_simulation_results.index))
            base_qty_series=pd.to_numeric(base_qty_series,errors='coerce').fillna(0)
            conditionnement_series_sim=df_simulation_results["Conditionnement"]
            adjusted_qty_series=(np.ceil(np.divide(base_qty_series,conditionnement_series_sim,out=np.zeros_like(base_qty_series,dtype=float),where=conditionnement_series_sim!=0))*conditionnement_series_sim).fillna(0).astype(int)
            df_simulation_results[forecast_qty_col_name]=adjusted_qty_series;df_simulation_results[fcst_amt_col]=adjusted_qty_series*df_simulation_results["Tarif d'achat"]
            total_adjusted_qty_all_months_series+=adjusted_qty_series;total_final_amount_all_months_series+=df_simulation_results[fcst_amt_col]
        df_simulation_results["Qté Totale Prév. (Mois Sel.)"]=total_adjusted_qty_all_months_series
        df_simulation_results["Montant Total Prév. (€) (Mois Sel.)"]=total_final_amount_all_months_series
        id_cols_display=["Fournisseur","Référence Article","Désignation Article","Conditionnement","Tarif d'achat"]
        n1_sales_cols_display=sorted([f"Ventes N-1 {m}"for m in selected_months_list if f"Ventes N-1 {m}"in df_simulation_results.columns])
        qty_forecast_cols_display=sorted([f"Qté Prév. {m}"for m in selected_months_list if f"Qté Prév. {m}"in df_simulation_results.columns])
        amt_forecast_cols_display=sorted([f"Montant Prév. {m} (€)"for m in selected_months_list if f"Montant Prév. {m} (€)"in df_simulation_results.columns])
        df_simulation_results.rename(columns={"Qté Totale Prév. (Mois Sel.)":"Qté Tot Prév (Mois Sel.)","Montant Total Prév. (€) (Mois Sel.)":"Mnt Tot Prév (€) (Mois Sel.)"},inplace=True)
        total_summary_cols_display=["Vts N-1 Tot (Mois Sel.)","Qté Tot Prév (Mois Sel.)","Mnt Tot Prév (€) (Mois Sel.)"]
        final_ordered_columns=id_cols_display+total_summary_cols_display+n1_sales_cols_display+qty_forecast_cols_display+amt_forecast_cols_display
        final_ordered_columns_existing=[c for c in final_ordered_columns if c in df_simulation_results.columns]
        grand_total_forecast_amount=total_final_amount_all_months_series.sum()
        return df_simulation_results[final_ordered_columns_existing],grand_total_forecast_amount
    except KeyError as e:st.error(f"Err clé (simu fcst): '{e}'.");logging.exception(f"KeyError in calc_fcst_sim_v3: {e}");return None,0.0
    except Exception as e:st.error(f"Err inattendue (simu fcst): {type(e).__name__} - {e}");logging.exception("Error in calc_fcst_sim_v3:");return None,0.0

def calculate_supplier_evaluation_and_targets(df_products, all_sales_cols, global_stock_target_value):
    supplier_data = {}
    if df_products.empty:
        st.warning("Aucune donnée produit pour l'évaluation des fournisseurs.")
        return supplier_data
    if "Tarif d'achat" not in df_products.columns:
        st.error("Colonne 'Tarif d'achat' manquante. Calcul du CA d'achat impossible.")
        return supplier_data
    df_eval = df_products.copy()
    df_eval["Tarif d'achat"] = pd.to_numeric(df_eval["Tarif d'achat"], errors='coerce').fillna(0)
    num_weeks_for_ca = 54; sales_cols_for_ca = []
    if len(all_sales_cols) >= num_weeks_for_ca: sales_cols_for_ca = all_sales_cols[-num_weeks_for_ca:]
    elif all_sales_cols: sales_cols_for_ca = all_sales_cols; st.caption(f"Moins de 54 sem. de ventes. CA achat fournisseur calculé sur {len(all_sales_cols)} sem.")
    else:
        st.warning("Aucune col. vente pour calcul CA achat fournisseur.")
        for supplier_name in df_eval["Fournisseur"].astype(str).unique(): supplier_data[supplier_name] = {'cogs_54w': 0, 'cogs_pct': 0, 'max_stock_target': 0}
        return supplier_data
    for col in sales_cols_for_ca:
        if col in df_eval.columns: df_eval[col] = pd.to_numeric(df_eval[col], errors='coerce').fillna(0)
        else: df_eval[col] = 0.0
    df_eval["Ventes_Unites_Periode_CA"] = df_eval[[c for c in sales_cols_for_ca if c in df_eval.columns]].sum(axis=1)
    df_eval["CA_Achat_Produit_Periode"] = df_eval["Ventes_Unites_Periode_CA"] * df_eval["Tarif d'achat"]
    supplier_cogs_total = df_eval.groupby("Fournisseur")["CA_Achat_Produit_Periode"].sum()
    if supplier_cogs_total.empty:
        st.warning("Impossible de calculer le CA d'achat par fournisseur.")
        for supplier_name in df_eval["Fournisseur"].astype(str).unique(): supplier_data[supplier_name] = {'cogs_54w': 0, 'cogs_pct': 0, 'max_stock_target': 0}
        return supplier_data
    global_cogs_total = supplier_cogs_total.sum()
    unique_suppliers_in_data = df_eval["Fournisseur"].astype(str).unique()
    if global_cogs_total > 0:
        for supplier_name in unique_suppliers_in_data:
            cogs_supplier = supplier_cogs_total.get(supplier_name, 0)
            cogs_percentage = (cogs_supplier / global_cogs_total) if global_cogs_total else 0
            max_stock_for_supplier = global_stock_target_value * cogs_percentage
            supplier_data[supplier_name] = {'cogs_54w': cogs_supplier, 'cogs_pct': cogs_percentage * 100, 'max_stock_target': max_stock_for_supplier}
    else:
        st.warning("CA d'Achat Global = 0. Objectif stock réparti équitablement.")
        num_suppliers = len(unique_suppliers_in_data)
        target_per_supplier_if_zero_cogs = global_stock_target_value / num_suppliers if num_suppliers > 0 else 0
        for supplier_name in unique_suppliers_in_data: supplier_data[supplier_name] = {'cogs_54w': 0, 'cogs_pct': 0, 'max_stock_target': target_per_supplier_if_zero_cogs}
    logging.info(f"Évaluation fournisseur (CA Achat) calculée: {len(supplier_data)} fournisseurs.")
    return supplier_data

def parse_week_column_to_date(col_name_str):
    if not isinstance(col_name_str, str): col_name_str = str(col_name_str)
    match_sw = re.match(r"(\d{4})[SW](\d{1,2})", col_name_str, re.IGNORECASE)
    match_plain = re.match(r"(\d{4})(\d{2})", col_name_str)
    year, week_num = None, None
    if match_sw: year, week_num = int(match_sw.group(1)), int(match_sw.group(2))
    elif match_plain:
        potential_year, potential_week = int(match_plain.group(1)), int(match_plain.group(2))
        if 1 <= potential_week <= 53 and 1900 < potential_year < 2200 : year, week_num = potential_year, potential_week
        else: return None
    else: return None
    if year and week_num and (1 <= week_num <= 53):
        try: date_str_iso = f"{year}-W{week_num:02}-1"; return pd.to_datetime(date_str_iso, format="%G-W%V-%u")
        except ValueError as e: logging.error(f"Err converting {year}W{week_num} from '{col_name_str}': {e}"); return None
    return None

def ai_calculate_order_quantities(df_products_for_ai, historical_semaine_cols, num_forecast_weeks,
                                  min_order_amount_for_subset=0.0, apply_special_rules=True,
                                  df_events_global=None):
    if not PROPHET_AVAILABLE: st.error("Librairie Prophet (IA) non installée."); return None, 0.0
    if df_products_for_ai.empty: st.info("Aucune donnée produit pour prévision IA."); return None, 0.0
    base_req_cols = ["Stock", "Conditionnement", "Tarif d'achat", "Référence Article"]
    missing_base = [c for c in base_req_cols if c not in df_products_for_ai.columns and c != "Référence Article"]
    if missing_base: st.error(f"Cols de base manquantes (calcul IA): {', '.join(missing_base)}"); return None, 0.0
    
    df_calc_ai = df_products_for_ai.copy()
    for col_op in ["Stock", "Conditionnement", "Tarif d'achat"]:
        if col_op in df_calc_ai.columns: df_calc_ai[col_op] = pd.to_numeric(df_calc_ai[col_op], errors='coerce').fillna(0)
        else: st.error(f"Colonne critique '{col_op}' manquante."); return None, 0.0
    df_calc_ai["Conditionnement"] = df_calc_ai["Conditionnement"].apply(lambda x: int(x) if x > 0 else 1)
    
    parsed_sales_dates = []
    valid_sales_cols_for_model = []
    for col_hist in historical_semaine_cols:
        parsed_dt_obj = parse_week_column_to_date(col_hist)
        if parsed_dt_obj: parsed_sales_dates.append({'date': parsed_dt_obj, 'col_name': col_hist}); valid_sales_cols_for_model.append(col_hist)
        else: logging.warning(f"Colonne '{col_hist}' ignorée pour IA (parsing date échoué).")
    if not parsed_sales_dates: st.error("Aucune colonne de ventes historiques interprétable comme date pour l'IA."); return None, 0.0
    parsed_sales_df_map = pd.DataFrame(parsed_sales_dates).sort_values(by='date').reset_index(drop=True)

    all_regressor_names = []
    if df_events_global is not None and not df_events_global.empty and \
       'TypeEvenement' in df_events_global.columns and \
       'ModeleImpact' in df_events_global.columns:
        # df_events_global['TypeEvenement'] est déjà str grâce au chargement
        regressors_df = df_events_global[df_events_global['ModeleImpact'].astype(str).str.lower() == 'regresseur']
        if not regressors_df.empty:
            all_regressor_names = regressors_df['TypeEvenement'].unique().tolist()
            # S'assurer que les noms sont des chaînes avant de les nettoyer
            all_regressor_names = [re.sub(r'\W+', '_', str(name)).lower() for name in all_regressor_names if pd.notna(name)]


    df_calc_ai["Qté Cmdée (IA)"] = 0; df_calc_ai["Forecast Ventes (IA)"] = 0.0
    num_prods = len(df_calc_ai); progress_bar_placeholder = st.empty()

    for i, (prod_idx, prod_row) in enumerate(df_calc_ai.iterrows()):
        progress_bar_placeholder.progress((i + 1) / num_prods, text=f"Prévision IA: Article {i+1}/{num_prods}")
        prod_ref_log = prod_row.get("Référence Article", f"Index {prod_idx}")
        # logging.info(f"DEBUG IA: Traitement produit {prod_ref_log}") # DEBUG

        prod_ts_hist = [{'ds': ps_row['date'], 'y': prod_row.get(ps_row['col_name'], np.nan)} for _, ps_row in parsed_sales_df_map.iterrows()]
        prod_ts_df_fit = pd.DataFrame(prod_ts_hist).dropna(subset=['ds'])

        if prod_ts_df_fit['y'].notna().sum() < 12:
            logging.warning(f"Produit {prod_ref_log}: <12 points ventes. Prévision IA ignorée.");
            df_calc_ai.loc[prod_idx, "Qté Cmdée (IA)"] = 0; df_calc_ai.loc[prod_idx, "Forecast Ventes (IA)"] = 0.0; continue

        product_specific_holidays_df = None
        product_regressors_to_add = {} 

        if df_events_global is not None and not df_events_global.empty and 'Référence Article' in df_events_global.columns:
            prod_events = df_events_global[df_events_global['Référence Article'] == prod_ref_log].copy()
            # if not prod_events.empty: # DEBUG
            #     logging.info(f"DEBUG IA: Événements trouvés pour {prod_ref_log}:\n{prod_events}")
            # else:
            #     logging.info(f"DEBUG IA: Aucun événement spécifique pour {prod_ref_log} dans l'onglet Spécial.")

            if not prod_events.empty:
                prod_holidays_list = []
                holiday_events_for_prod = prod_events[prod_events['ModeleImpact'].astype(str).str.lower() == 'holiday']
                for _, event_row in holiday_events_for_prod.iterrows():
                    holiday_name_clean = re.sub(r'\W+', '_', str(event_row['TypeEvenement'])).lower()
                    holiday_entry = {'holiday': holiday_name_clean, 'ds': event_row['DateDebut']}
                    if pd.notna(event_row['DateFin']) and event_row['DateFin'] > event_row['DateDebut']:
                        date_range = pd.date_range(start=event_row['DateDebut'], end=event_row['DateFin'], freq='D')
                        for date_in_range in date_range:
                            prod_holidays_list.append({'holiday': holiday_name_clean, 'ds': date_in_range})
                    else:
                        prod_holidays_list.append(holiday_entry)
                if prod_holidays_list:
                    product_specific_holidays_df = pd.DataFrame(prod_holidays_list)
                    # logging.info(f"DEBUG IA: Holidays DF pour {prod_ref_log}:\n{product_specific_holidays_df}") # DEBUG

                for reg_name_sanitized in all_regressor_names:
                    original_type_event_for_reg = None
                    # Chercher le TypeEvenement original qui correspond au nom sanitizé, pour CE produit
                    relevant_prod_regressor_events = prod_events[prod_events['ModeleImpact'].astype(str).str.lower() == 'regresseur']
                    for _, ev_row_chk in relevant_prod_regressor_events.iterrows():
                        if re.sub(r'\W+', '_', str(ev_row_chk['TypeEvenement'])).lower() == reg_name_sanitized:
                            original_type_event_for_reg = ev_row_chk['TypeEvenement']
                            break
                    
                    default_reg_val = 0.0 # Valeur neutre pour additif
                    is_multiplicative_reg = False

                    if original_type_event_for_reg: # Si ce régresseur est défini pour ce produit
                        event_reg_data_for_prod = relevant_prod_regressor_events[
                            relevant_prod_regressor_events['TypeEvenement'] == original_type_event_for_reg
                        ]
                        if not event_reg_data_for_prod.empty and 'UniteImpact' in event_reg_data_for_prod.columns and \
                           str(event_reg_data_for_prod['UniteImpact'].iloc[0]).lower() == 'multiplicatif':
                            default_reg_val = 1.0 # Valeur neutre pour multiplicatif
                            is_multiplicative_reg = True
                        
                        prod_ts_df_fit[reg_name_sanitized] = default_reg_val 
                        product_regressors_to_add[reg_name_sanitized] = True 

                        for _, reg_row in event_reg_data_for_prod.iterrows():
                            start_d = reg_row['DateDebut']
                            end_d = reg_row['DateFin'] if pd.notna(reg_row['DateFin']) else prod_ts_df_fit['ds'].max() + pd.Timedelta(days=730) 
                            val_impact = reg_row['ValeurImpact']
                            unit_impact = str(reg_row.get('UniteImpact', '')).lower()

                            if pd.isna(val_impact): continue
                            
                            mask = (prod_ts_df_fit['ds'] >= start_d) & (prod_ts_df_fit['ds'] <= end_d)
                            if unit_impact == 'additifhebdo': prod_ts_df_fit.loc[mask, reg_name_sanitized] = val_impact
                            elif unit_impact == 'multiplicatif': prod_ts_df_fit.loc[mask, reg_name_sanitized] = val_impact
                            else: prod_ts_df_fit.loc[mask, reg_name_sanitized] = val_impact
                    else: # Ce régresseur global n'est pas spécifiquement défini pour ce produit
                          # On doit quand même créer la colonne avec une valeur neutre si d'autres produits l'utilisent
                        if reg_name_sanitized not in prod_ts_df_fit.columns: # S'assurer qu'elle n'existe pas déjà
                             # Déterminer si ce régresseur (globalement) est multiplicatif pour la valeur neutre
                            temp_event_for_global_mode = df_events_global[
                                (df_events_global['TypeEvenement'].apply(lambda x: re.sub(r'\W+', '_', str(x)).lower()) == reg_name_sanitized) &
                                (df_events_global['ModeleImpact'].astype(str).str.lower() == 'regresseur')
                            ]
                            if not temp_event_for_global_mode.empty and 'UniteImpact' in temp_event_for_global_mode.columns and \
                               str(temp_event_for_global_mode['UniteImpact'].iloc[0]).lower() == 'multiplicatif':
                                prod_ts_df_fit[reg_name_sanitized] = 1.0
                            else:
                                prod_ts_df_fit[reg_name_sanitized] = 0.0
        try:
            model_prophet = Prophet(
                uncertainty_samples=0,
                holidays=product_specific_holidays_df
            )

            active_regressors_for_this_model = []
            for reg_name_to_add in all_regressor_names:
                 if product_regressors_to_add.get(reg_name_to_add): # Si ce régresseur a été marqué comme pertinent
                    mode = 'additive'
                    original_type_for_mode = None
                    # Retrouver le type original pour ce nom de regresseur sanitizé
                    temp_prod_events_reg = prod_events[prod_events['ModeleImpact'].astype(str).str.lower() == 'regresseur']
                    for _, ev_row_mode in temp_prod_events_reg.iterrows():
                        if re.sub(r'\W+', '_', str(ev_row_mode['TypeEvenement'])).lower() == reg_name_to_add:
                            original_type_for_mode = ev_row_mode['TypeEvenement']
                            break
                    
                    if original_type_for_mode:
                        event_for_mode_check = temp_prod_events_reg[temp_prod_events_reg['TypeEvenement'] == original_type_for_mode]
                        if not event_for_mode_check.empty and 'UniteImpact' in event_for_mode_check.columns:
                            if str(event_for_mode_check['UniteImpact'].iloc[0]).lower() == 'multiplicatif':
                                mode = 'multiplicative'
                    
                    model_prophet.add_regressor(reg_name_to_add, mode=mode)
                    active_regressors_for_this_model.append(reg_name_to_add)


            if not prod_ts_df_fit.empty and (prod_ts_df_fit['ds'].max() - prod_ts_df_fit['ds'].min()) >= pd.Timedelta(days=365 + 180):
                model_prophet.add_seasonality(name='yearly', period=365.25, fourier_order=10)
            
            with SuppressStdoutStderr():
                cols_for_fit = ['ds', 'y'] + active_regressors_for_this_model
                # S'assurer que toutes les colonnes actives sont bien dans prod_ts_df_fit
                for col_fit_check in active_regressors_for_this_model:
                    if col_fit_check not in prod_ts_df_fit.columns:
                        # Ceci ne devrait pas arriver si la logique précédente est correcte
                        logging.warning(f"Regresseur {col_fit_check} actif mais non trouvé dans prod_ts_df_fit pour {prod_ref_log}. Ajout avec valeur neutre.")
                        # Redéterminer la valeur neutre
                        is_mult = False
                        orig_type_neut_fit = None
                        temp_prod_ev_neut_fit = prod_events[prod_events['ModeleImpact'].astype(str).str.lower() == 'regresseur']
                        for _, ev_row_neut_fit in temp_prod_ev_neut_fit.iterrows():
                            if re.sub(r'\W+', '_', str(ev_row_neut_fit['TypeEvenement'])).lower() == col_fit_check:
                                orig_type_neut_fit = ev_row_neut_fit['TypeEvenement']
                                break
                        if orig_type_neut_fit:
                            ev_check_neut_fit = temp_prod_ev_neut_fit[temp_prod_ev_neut_fit['TypeEvenement'] == orig_type_neut_fit]
                            if not ev_check_neut_fit.empty and 'UniteImpact' in ev_check_neut_fit.columns and \
                               str(ev_check_neut_fit['UniteImpact'].iloc[0]).lower() == 'multiplicatif':
                                is_mult = True
                        prod_ts_df_fit[col_fit_check] = 1.0 if is_mult else 0.0


                df_fit_final = prod_ts_df_fit[cols_for_fit].copy() # Utiliser .copy()
                df_fit_final.dropna(subset=['y'], inplace=True) # Garder les y non-NA, les régresseurs NA seront 0/1
                for r_col in active_regressors_for_this_model: # Remplir les NA restants dans les régresseurs
                    is_mult_fill = False
                    orig_type_fill = None
                    temp_prod_ev_fill = prod_events[prod_events['ModeleImpact'].astype(str).str.lower() == 'regresseur']
                    for _, ev_row_fill in temp_prod_ev_fill.iterrows():
                        if re.sub(r'\W+', '_', str(ev_row_fill['TypeEvenement'])).lower() == r_col:
                            orig_type_fill = ev_row_fill['TypeEvenement']
                            break
                    if orig_type_fill:
                        ev_check_fill = temp_prod_ev_fill[temp_prod_ev_fill['TypeEvenement'] == orig_type_fill]
                        if not ev_check_fill.empty and 'UniteImpact' in ev_check_fill.columns and \
                           str(ev_check_fill['UniteImpact'].iloc[0]).lower() == 'multiplicatif':
                            is_mult_fill = True
                    df_fit_final[r_col].fillna(1.0 if is_mult_fill else 0.0, inplace=True)
                
                # logging.info(f"DEBUG IA: df_fit_final pour {prod_ref_log} (avant fit):\n{df_fit_final}") # DEBUG
                model_prophet.fit(df_fit_final)

            future_df = model_prophet.make_future_dataframe(periods=num_forecast_weeks, freq='W-MON')

            if df_events_global is not None and not df_events_global.empty and active_regressors_for_this_model:
                for reg_name_fut in active_regressors_for_this_model: # Seulement les régresseurs actifs pour CE modèle
                    original_type_event_for_fut = None
                    temp_prod_ev_fut = prod_events[prod_events['ModeleImpact'].astype(str).str.lower() == 'regresseur']
                    for _, ev_row_fut_find in temp_prod_ev_fut.iterrows():
                        if re.sub(r'\W+', '_', str(ev_row_fut_find['TypeEvenement'])).lower() == reg_name_fut:
                            original_type_event_for_fut = ev_row_fut_find['TypeEvenement']
                            break
                    
                    event_reg_fut_data = pd.DataFrame()
                    if original_type_event_for_fut:
                        event_reg_fut_data = temp_prod_ev_fut[temp_prod_ev_fut['TypeEvenement'] == original_type_event_for_fut]
                    
                    default_reg_val_fut = 0.0
                    if not event_reg_fut_data.empty and 'UniteImpact' in event_reg_fut_data.columns and \
                        str(event_reg_fut_data['UniteImpact'].iloc[0]).lower() == 'multiplicatif':
                        default_reg_val_fut = 1.0
                    future_df[reg_name_fut] = default_reg_val_fut

                    if not event_reg_fut_data.empty:
                        for _, reg_row_fut in event_reg_fut_data.iterrows():
                            start_d_fut = reg_row_fut['DateDebut']
                            end_d_fut = reg_row_fut['DateFin'] if pd.notna(reg_row_fut['DateFin']) else future_df['ds'].max()
                            val_impact_fut = reg_row_fut['ValeurImpact']
                            unit_impact_fut = str(reg_row_fut.get('UniteImpact', '')).lower()

                            if pd.isna(val_impact_fut): continue

                            mask_fut = (future_df['ds'] >= start_d_fut) & (future_df['ds'] <= end_d_fut)
                            if unit_impact_fut == 'additifhebdo': future_df.loc[mask_fut, reg_name_fut] = val_impact_fut
                            elif unit_impact_fut == 'multiplicatif': future_df.loc[mask_fut, reg_name_fut] = val_impact_fut
                            else: future_df.loc[mask_fut, reg_name_fut] = val_impact_fut
            
            # S'assurer que toutes les colonnes de régresseurs attendues par le modèle sont dans future_df
            for regressor_name_in_model in model_prophet.extra_regressors.keys():
                if regressor_name_in_model not in future_df.columns:
                    logging.warning(f"Regresseur {regressor_name_in_model} attendu par le modèle mais non trouvé dans future_df pour {prod_ref_log}. Ajout avec valeur neutre.")
                    # Déterminer la valeur neutre
                    is_mult_fut_fill = False
                    orig_type_fut_fill = None
                    temp_prod_ev_fut_fill = prod_events[prod_events['ModeleImpact'].astype(str).str.lower() == 'regresseur']
                    for _, ev_row_fut_fill in temp_prod_ev_fut_fill.iterrows():
                        if re.sub(r'\W+', '_', str(ev_row_fut_fill['TypeEvenement'])).lower() == regressor_name_in_model:
                            orig_type_fut_fill = ev_row_fut_fill['TypeEvenement']
                            break
                    if orig_type_fut_fill:
                        ev_check_fut_fill = temp_prod_ev_fut_fill[temp_prod_ev_fut_fill['TypeEvenement'] == orig_type_fut_fill]
                        if not ev_check_fut_fill.empty and 'UniteImpact' in ev_check_fut_fill.columns and \
                           str(ev_check_fut_fill['UniteImpact'].iloc[0]).lower() == 'multiplicatif':
                            is_mult_fut_fill = True
                    future_df[regressor_name_in_model] = 1.0 if is_mult_fut_fill else 0.0

            # logging.info(f"DEBUG IA: future_df pour {prod_ref_log} (avant predict):\n{future_df[ ['ds'] + list(model_prophet.extra_regressors.keys()) ].tail()}") # DEBUG
            forecast_df_res = model_prophet.predict(future_df)
            total_fcst_period = forecast_df_res['yhat'].iloc[-num_forecast_weeks:].sum()
            total_fcst_period = max(0, total_fcst_period) # Assurer que le forecast n'est pas négatif
            df_calc_ai.loc[prod_idx, "Forecast Ventes (IA)"] = total_fcst_period

            stock_item = prod_row["Stock"]; package_item = prod_row["Conditionnement"]
            needed_raw = total_fcst_period - stock_item
            order_qty_item_ia = 0
            if needed_raw > 0:
                if package_item > 0: order_qty_item_ia = int(np.ceil(needed_raw / package_item) * package_item)
                else: logging.warning(f"Produit {prod_ref_log}: Cond. {package_item} invalide. Cmd IA=0.")
            
            if apply_special_rules and order_qty_item_ia == 0 and stock_item <= 1 and package_item > 0:
                recent_sales_cols_chk = [psc_row['col_name'] for psc_row in parsed_sales_df_map.tail(12).to_dict('records')]
                actual_recent_cols = [c for c in recent_sales_cols_chk if c in df_calc_ai.columns] # Assurer que les colonnes existent dans df_calc_ai
                if actual_recent_cols and df_calc_ai.loc[prod_idx, actual_recent_cols].sum(skipna=True) > 0: # Utiliser skipna=True
                    order_qty_item_ia = package_item
                    logging.info(f"Produit {prod_ref_log}: Stock bas, vts récentes ({df_calc_ai.loc[prod_idx, actual_recent_cols].sum(skipna=True)}), fcst IA=0. Forçage à 1 cond ({package_item}).")
            
            df_calc_ai.loc[prod_idx, "Qté Cmdée (IA)"] = order_qty_item_ia
        except Exception as e_ph: 
            logging.exception(f"Erreur Prophet pour {prod_ref_log}: {e_ph}")
            st.error(f"Erreur IA (produit {prod_ref_log}): {e_ph}. Vérifiez config. événements ou données produit.")
            df_calc_ai.loc[prod_idx, "Qté Cmdée (IA)"] = 0
            df_calc_ai.loc[prod_idx, "Forecast Ventes (IA)"] = 0.0
    
    progress_bar_placeholder.empty()
    df_calc_ai["Total Cmd (€) (IA)"] = df_calc_ai["Qté Cmdée (IA)"] * df_calc_ai["Tarif d'achat"]
    current_total_amount_ia = df_calc_ai["Total Cmd (€) (IA)"].sum()
    if min_order_amount_for_subset > 0 and current_total_amount_ia < min_order_amount_for_subset:
        logging.info(f"Ajustement IA pour min cmd: {min_order_amount_for_subset:,.2f}€. Actuel: {current_total_amount_ia:,.2f}€")
        eligible_inc_indices = df_calc_ai[(df_calc_ai["Qté Cmdée (IA)"] > 0) & (df_calc_ai["Conditionnement"] > 0) & (df_calc_ai["Tarif d'achat"] > 0)].index.tolist()
        if not eligible_inc_indices: st.warning(f"Min cmd (IA) de {min_order_amount_for_subset:,.2f}€ non atteint. Aucun article éligible.")
        else:
            item_ptr_adj = 0; max_adj_iter = len(eligible_inc_indices) * 20 + 1; current_adj_iter = 0
            qtes_cmdees_ia_series_adj = df_calc_ai["Qté Cmdée (IA)"].copy()
            while current_total_amount_ia < min_order_amount_for_subset and current_adj_iter < max_adj_iter:
                current_adj_iter += 1
                df_item_idx_inc = eligible_inc_indices[item_ptr_adj]
                pkg_adj = df_calc_ai.loc[df_item_idx_inc, "Conditionnement"]; price_adj = df_calc_ai.loc[df_item_idx_inc, "Tarif d'achat"]
                if pkg_adj > 0 and price_adj > 0: qtes_cmdees_ia_series_adj.loc[df_item_idx_inc] += pkg_adj; current_total_amount_ia += (pkg_adj * price_adj)
                else: logging.warning(f"Skipping min order increment for item index {df_item_idx_inc} due to invalid pkg/price.")
                item_ptr_adj = (item_ptr_adj + 1) % len(eligible_inc_indices)
            df_calc_ai["Qté Cmdée (IA)"] = qtes_cmdees_ia_series_adj
            if current_adj_iter >= max_adj_iter and current_total_amount_ia < min_order_amount_for_subset: st.error(f"Ajustement min (IA): Max itérations. Actuel: {current_total_amount_ia:,.2f}€ / Requis: {min_order_amount_for_subset:,.2f}€.")
            else: logging.info(f"Montant après ajustement IA pour min: {current_total_amount_ia:,.2f}€")
            df_calc_ai["Total Cmd (€) (IA)"] = df_calc_ai["Qté Cmdée (IA)"] * df_calc_ai["Tarif d'achat"]
            current_total_amount_ia = df_calc_ai["Total Cmd (€) (IA)"].sum()
    df_calc_ai["Stock Terme (IA)"] = df_calc_ai["Stock"] + df_calc_ai["Qté Cmdée (IA)"]
    return df_calc_ai, current_total_amount_ia

# --- Streamlit App UI ---
st.set_page_config(page_title="Forecast & Rotation App", layout="wide")
st.title("📦 Application Prévision Commande, Analyse Rotation & Suivi")
uploaded_file = st.file_uploader("📁 Charger le fichier Excel principal", type=["xlsx", "xls"], key="main_file_uploader")

# --- DEBUG SECTION (Optionnel, à activer pour le débogage) ---
# show_debug_info = st.sidebar.checkbox("Afficher infos de débogage onglet Spécial")
# if show_debug_info:
#     if 'df_product_events' in st.session_state and not st.session_state.df_product_events.empty:
#         st.sidebar.subheader("DEBUG: Contenu df_product_events")
#         st.sidebar.dataframe(st.session_state.df_product_events.head(10)) # Afficher les 10 premières lignes
#         st.sidebar.markdown(f"Nombre total d'événements chargés: {len(st.session_state.df_product_events)}")
#         st.sidebar.markdown(f"Colonnes: {st.session_state.df_product_events.columns.tolist()}")
#         st.sidebar.markdown("Infos DataFrame:")
#         buf = io.StringIO()
#         st.session_state.df_product_events.info(buf=buf)
#         st.sidebar.text(buf.getvalue())
#     elif 'df_product_events' in st.session_state:
#         st.sidebar.warning("DEBUG: df_product_events est VIDE.")
#     else:
#         st.sidebar.error("DEBUG: df_product_events N'EXISTE PAS dans session_state.")
# --- FIN DEBUG SECTION ---


def get_default_session_state():
    return {
        'df_full': None, 'min_order_dict': {}, 'df_initial_filtered': pd.DataFrame(),
        'all_available_semaine_columns': [], 'unique_suppliers_list': [],
        'commande_result_df': None, 'commande_calculated_total_amount': 0.0,
        'commande_suppliers_calculated_for': [], 'commande_params_calculated_for': {},
        'ai_commande_result_df': None, 'ai_commande_total_amount': 0.0,
        'ai_commande_params_calculated_for': {}, 'ai_forecast_weeks_val': 4, 'ai_min_order_val': 0.0,
        'ai_ignored_orders_df': None,
        'ai_excluded_suppliers_stock_target': [], 
        'df_product_events': pd.DataFrame(), 
        'supplier_evaluation_data': None, 'global_stock_target_config': 3200000.0,
        'rotation_result_df': None, 'rotation_analysis_period_label': "12 dernières semaines",
        'rotation_suppliers_calculated_for': [], 'rotation_threshold_value': 1.0,
        'show_all_rotation_data': True, 'rotation_params_calculated_for': {},
        'forecast_result_df': None, 'forecast_grand_total_amount': 0.0,
        'forecast_simulation_params_calculated_for': {},
        'forecast_selected_months_ui': list(calendar.month_name)[1:],
        'forecast_sim_type_radio_index': 0, 'forecast_progression_percentage_ui': 5.0,
        'forecast_target_amount_ui': 10000.0,
        'df_suivi_commandes': pd.DataFrame(),
    }

for key, default_value in get_default_session_state().items():
    if key not in st.session_state: st.session_state[key] = default_value

if uploaded_file and st.session_state.df_full is None:
    logging.info(f"Nouveau fichier: {uploaded_file.name}. Réinitialisation...")
    dynamic_prefixes = ['tab1_', 'tab1_ai_', 'tab2_', 'tab4_', 'tab5_']
    keys_to_del_from_session = [k for k in st.session_state if k in get_default_session_state() or any(k.startswith(p) for p in dynamic_prefixes)]
    for k_del in keys_to_del_from_session:
        try: del st.session_state[k_del]
        except KeyError: pass
    for key_init, val_init in get_default_session_state().items():
        st.session_state[key_init] = val_init
    logging.info("État session réinitialisé.")

    try:
        excel_io_buf = io.BytesIO(uploaded_file.getvalue())
        st.info("Lecture 'Tableau final'...")
        df_full_read = safe_read_excel(excel_io_buf, sheet_name="Tableau final", header=7)
        if df_full_read is None or df_full_read.empty: st.error("❌ Échec lecture 'Tableau final' ou onglet vide."); st.stop()
        
        req_tf_cols_check = ["Stock", "Fournisseur", "AF_RefFourniss", "Tarif d'achat", "Conditionnement", "Référence Article", "Désignation Article", "Date Création Article"]
        missing_tf_check = [c for c in req_tf_cols_check if c not in df_full_read.columns]
        if missing_tf_check: st.error(f"❌ Cols manquantes ('TF'): {', '.join(missing_tf_check)}. Vérifiez ligne en-tête (L8)."); st.stop()
        
        df_full_read["Stock"] = pd.to_numeric(df_full_read["Stock"], errors='coerce').fillna(0)
        df_full_read["Tarif d'achat"] = pd.to_numeric(df_full_read["Tarif d'achat"], errors='coerce').fillna(0)
        df_full_read["Conditionnement"] = pd.to_numeric(df_full_read["Conditionnement"], errors='coerce').fillna(1).apply(lambda x: int(x) if x > 0 else 1)
        
        if "Date Création Article" in df_full_read.columns:
            try:
                df_full_read["Date Création Article"] = pd.to_datetime(df_full_read["Date Création Article"], errors='coerce')
                if df_full_read["Date Création Article"].isnull().any():
                    st.warning("⚠️ Certaines dates de création d'article n'ont pas pu être lues et seront ignorées pour la recherche de nouveaux articles.")
            except Exception as e_date_creation:
                st.error(f"❌ Erreur conversion 'Date Création Article': {e_date_creation}.")
        else: 
             st.warning("⚠️ Colonne 'Date Création Article' non trouvée. Recherche nouveaux articles désactivée.")

        for str_c_tf in ["Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article"]:
            if str_c_tf in df_full_read.columns: df_full_read[str_c_tf] = df_full_read[str_c_tf].astype(str).str.strip().replace('nan', '')
        st.session_state.df_full = df_full_read
        st.success("✅ 'TF' lu.")

        st.info("Lecture 'Min commande'...")
        excel_io_buf.seek(0)
        df_min_c_read = safe_read_excel(excel_io_buf, sheet_name="Minimum de commande")
        min_o_dict_temp_read = {}
        if df_min_c_read is not None and not df_min_c_read.empty:
            s_col_min, m_col_min = "Fournisseur", "Minimum de Commande"
            if s_col_min in df_min_c_read.columns and m_col_min in df_min_c_read.columns:
                try:
                    df_min_c_read[s_col_min] = df_min_c_read[s_col_min].astype(str).str.strip().replace('nan', '')
                    df_min_c_read[m_col_min] = pd.to_numeric(df_min_c_read[m_col_min], errors='coerce')
                    min_o_dict_temp_read = df_min_c_read.dropna(subset=[s_col_min, m_col_min]).set_index(s_col_min)[m_col_min].to_dict()
                    st.success(f"✅ 'Min cmd' lu ({len(min_o_dict_temp_read)} entrées).")
                except Exception as e_min_proc: st.error(f"❌ Err trait. 'Min cmd': {e_min_proc}")
            else: st.warning(f"⚠️ Cols '{s_col_min}'/'{m_col_min}' manquantes ('Min cmd').")
        elif df_min_c_read is None: st.info("Onglet 'Min cmd' non trouvé.")
        else: st.info("Onglet 'Min cmd' vide.")
        st.session_state.min_order_dict = min_o_dict_temp_read

        st.info("Lecture 'Suivi commandes'...")
        excel_io_buf.seek(0)
        df_suivi_read = safe_read_excel(excel_io_buf, sheet_name="Suivi commandes", header=4)
        if df_suivi_read is not None and not df_suivi_read.empty:
            req_s_cols_check = ["Date Pièce BC", "N° de pièce", "AF_RefFourniss", "Désignation Article", "Qté Commandées", "Intitulé Fournisseur"]
            miss_s_cols_c_check = [c for c in req_s_cols_check if c not in df_suivi_read.columns]
            if not miss_s_cols_c_check:
                df_suivi_read.rename(columns={"Intitulé Fournisseur": "Fournisseur"}, inplace=True)
                for col_strp_s in ["Fournisseur", "AF_RefFourniss", "Désignation Article", "N° de pièce"]:
                    if col_strp_s in df_suivi_read.columns: df_suivi_read[col_strp_s] = df_suivi_read[col_strp_s].astype(str).str.strip().replace('nan','')
                if "Qté Commandées" in df_suivi_read.columns: df_suivi_read["Qté Commandées"] = pd.to_numeric(df_suivi_read["Qté Commandées"], errors='coerce').fillna(0)
                if "Date Pièce BC" in df_suivi_read.columns:
                    try: df_suivi_read["Date Pièce BC"] = pd.to_datetime(df_suivi_read["Date Pièce BC"], errors='coerce')
                    except Exception as e_dt_s: st.warning(f"⚠️ Problème parsing 'Date Pièce BC' (Suivi): {e_dt_s}.")
                df_suivi_read.dropna(how='all', inplace=True)
                st.session_state.df_suivi_commandes = df_suivi_read
                st.success(f"✅ 'Suivi cmds' lu ({len(df_suivi_read)} lignes).")
            else:
                st.warning(f"⚠️ Cols manquantes ('Suivi cmds', L5): {', '.join(miss_s_cols_c_check)}. Suivi limité.")
                st.session_state.df_suivi_commandes = pd.DataFrame()
        elif df_suivi_read is None: st.info("Onglet 'Suivi cmds' non trouvé.")
        else: st.info("Onglet 'Suivi cmds' vide."); st.session_state.df_suivi_commandes = pd.DataFrame()

        st.info("Lecture onglet 'Spécial' pour paramètres IA avancés...")
        excel_io_buf.seek(0) 
        df_events_temp = safe_read_excel(excel_io_buf, sheet_name="Spécial")

        if df_events_temp is not None and not df_events_temp.empty:
            df_events_temp.columns = df_events_temp.columns.str.strip()
            required_event_cols = ['Référence Article', 'TypeEvenement', 'DateDebut']
            
            missing_cols_special = [col for col in required_event_cols if col not in df_events_temp.columns]
            if not missing_cols_special:
                try:
                    df_events_temp['DateDebut'] = pd.to_datetime(df_events_temp['DateDebut'], errors='coerce')
                    
                    if 'DateFin' in df_events_temp.columns: df_events_temp['DateFin'] = pd.to_datetime(df_events_temp['DateFin'], errors='coerce')
                    else: df_events_temp['DateFin'] = pd.NaT 

                    if 'ValeurImpact' in df_events_temp.columns: df_events_temp['ValeurImpact'] = pd.to_numeric(df_events_temp['ValeurImpact'], errors='coerce')
                    else: df_events_temp['ValeurImpact'] = np.nan

                    if 'UniteImpact' not in df_events_temp.columns: df_events_temp['UniteImpact'] = 'multiplicatif' # minuscule par défaut
                    else: df_events_temp['UniteImpact'] = df_events_temp['UniteImpact'].astype(str).str.lower()
                    
                    if 'ModeleImpact' not in df_events_temp.columns: df_events_temp['ModeleImpact'] = 'regresseur' # minuscule par défaut
                    else: df_events_temp['ModeleImpact'] = df_events_temp['ModeleImpact'].astype(str).str.lower()

                    # S'assurer que les colonnes pour dropna existent
                    cols_for_dropna = ['Référence Article', 'TypeEvenement', 'DateDebut', 'ModeleImpact']
                    for col_dn in cols_for_dropna:
                        if col_dn not in df_events_temp.columns: # Si une colonne critique manque après les valeurs par défaut
                            st.warning(f"Colonne critique '{col_dn}' manquante dans l'onglet 'Spécial' même après tentative de valeur par défaut. Impossible de traiter les événements.")
                            df_events_temp = pd.DataFrame() # Invalider pour ne pas continuer
                            break
                    
                    df_events_processed = pd.DataFrame() # Initialiser
                    if not df_events_temp.empty: # Vérifier si df_events_temp n'a pas été invalidé
                        df_events_processed = df_events_temp.dropna(subset=cols_for_dropna).copy()
                    
                        if 'TypeEvenement' in df_events_processed.columns:
                            df_events_processed['TypeEvenement'] = df_events_processed['TypeEvenement'].astype(str)
                    
                    st.session_state.df_product_events = df_events_processed
                    if not df_events_processed.empty:
                        st.success(f"✅ Onglet 'Spécial' lu ({len(st.session_state.df_product_events)} événements/paramètres IA).")
                    elif not df_events_temp.empty: # df_events_temp n'était pas vide, mais après dropna, df_events_processed l'est
                        st.warning("Onglet 'Spécial' lu, mais aucune ligne valide trouvée après filtrage (vérifiez les valeurs manquantes dans les colonnes clés).")
                    # Si df_events_temp était déjà vide (pris en charge par la condition externe)

                except Exception as e:
                    st.warning(f"⚠️ Erreur traitement onglet 'Spécial': {e}. Paramètres IA avancés ignorés.")
                    st.session_state.df_product_events = pd.DataFrame()
            else:
                st.warning(f"⚠️ Onglet 'Spécial' trouvé, mais colonnes requises manquantes: {', '.join(missing_cols_special)}. Paramètres IA avancés ignorés.")
                st.session_state.df_product_events = pd.DataFrame()
        elif df_events_temp is None: 
            st.info("ℹ️ Onglet 'Spécial' non trouvé. Aucuns paramètres IA avancés chargés.")
            st.session_state.df_product_events = pd.DataFrame() 
        else: 
            st.info("ℹ️ Onglet 'Spécial' est vide. Aucuns paramètres IA avancés chargés.")
            st.session_state.df_product_events = pd.DataFrame()

        df_full_state = st.session_state.df_full
        df_init_filt_temp_read = df_full_state[
            (df_full_state["Fournisseur"].astype(str).str.strip() != "") &
            (df_full_state["Fournisseur"].astype(str).str.strip().str.lower() != "#filter") &
            (df_full_state["AF_RefFourniss"].astype(str).str.strip() != "")
        ].copy()
        st.session_state.df_initial_filtered = df_init_filt_temp_read

        first_week_col_idx_approx = 12
        potential_sem_cols_read = []
        if len(df_full_state.columns) > first_week_col_idx_approx:
            candidate_cols_sem = df_full_state.columns[first_week_col_idx_approx:].tolist()
            known_non_week_cols_set = set(["Tarif d'achat", "Conditionnement", "Stock", "Total", "Stock à terme", "Ventes N-1", "Ventes 12 semaines identiques N-1", "Ventes 12 dernières semaines", "Quantité à commander", "Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article", "Date Création Article"])
            for col_cand_sem in candidate_cols_sem:
                if col_cand_sem not in known_non_week_cols_set:
                    try:
                        is_numeric_like = pd.to_numeric(df_full_state[col_cand_sem], errors='coerce').notna().sum() > (len(df_full_state) * 0.1)
                        is_date_col_name = parse_week_column_to_date(str(col_cand_sem)) is not None
                        if is_numeric_like or is_date_col_name:
                            potential_sem_cols_read.append(col_cand_sem)
                    except Exception: pass
        st.session_state.all_available_semaine_columns = potential_sem_cols_read
        if not potential_sem_cols_read: st.warning("⚠️ Aucune col vente numérique/datable auto-identifiée après la 12ème. Vérifiez le format.")

        if not df_init_filt_temp_read.empty:
            st.session_state.unique_suppliers_list = sorted(df_init_filt_temp_read["Fournisseur"].astype(str).unique().tolist())
            st.session_state.supplier_evaluation_data = calculate_supplier_evaluation_and_targets(
                st.session_state.df_initial_filtered,
                st.session_state.all_available_semaine_columns,
                st.session_state.global_stock_target_config
            )
            if st.session_state.supplier_evaluation_data:
                st.success("✅ Évaluation CA fournisseur et objectifs stock calculés.")
        else:
            st.session_state.unique_suppliers_list = []
            st.session_state.supplier_evaluation_data = {}

        st.success("✅ Fichier principal chargé et données initiales préparées.")
        st.rerun()
    except Exception as e_load_main_fatal:
        st.error(f"❌ Err majeure chargement/traitement: {e_load_main_fatal}")
        logging.exception("Major file loading/processing error:")
        st.session_state.df_full = None; st.session_state.df_initial_filtered = pd.DataFrame()
        st.stop()

# --- Main App UI ---
if 'df_initial_filtered' in st.session_state and isinstance(st.session_state.df_initial_filtered, pd.DataFrame):
    df_base_tabs = st.session_state.df_initial_filtered
    all_sups_data = st.session_state.unique_suppliers_list
    min_o_amts = st.session_state.min_order_dict
    id_sem_cols = st.session_state.all_available_semaine_columns
    df_suivi_cmds_all = st.session_state.get('df_suivi_commandes', pd.DataFrame())

    tab_titles_main = ["Prévision Commande", "Prévision Commande (IA)", "Analyse Rotation Stock",
                       "Vérification Stock", "Simulation Forecast", "Suivi Commandes Fourn.", "Nouveaux Articles"]
    tab1, tab1_ai, tab2, tab3, tab4, tab5, tab6 = st.tabs(tab_titles_main)

    # --- Tab 1: Classic Order Forecast ---
    with tab1:
        # ... (Code identique à la version précédente)
        st.header("Prévision des Quantités à Commander (Méthode Classique)")
        sel_f_t1 = render_supplier_checkboxes("tab1", all_sups_data, default_select_all=True)
        df_disp_t1 = pd.DataFrame()
        if sel_f_t1:
            if not df_base_tabs.empty: df_disp_t1 = df_base_tabs[df_base_tabs["Fournisseur"].isin(sel_f_t1)].copy(); st.caption(f"{len(df_disp_t1)} art. / {len(sel_f_t1)} fourn.")
        else:st.info("Sélectionner fournisseur(s).")
        st.markdown("---")
        if df_disp_t1.empty and sel_f_t1:st.warning("Aucun article pour fournisseur(s) sélectionné(s).")
        elif not id_sem_cols and not df_disp_t1.empty:st.warning("Colonnes ventes non identifiées.")
        elif not df_disp_t1.empty:
            st.markdown("#### Paramètres Calcul Commande")
            c1_c,c2_c=st.columns(2);
            default_duree_t1 = st.session_state.get('commande_params_calculated_for',{}).get('duree_semaines', 4)
            default_min_amt_t1 = st.session_state.get('commande_params_calculated_for',{}).get('min_amount', 0.0)
            if len(sel_f_t1) == 1 and sel_f_t1[0] in min_o_amts and default_min_amt_t1 == 0.0:
                default_min_amt_t1 = min_o_amts[sel_f_t1[0]]

            with c1_c:d_s_c_t1=st.number_input("⏳ Couverture (sem.)",1,260,value=default_duree_t1,step=1,key="d_s_c_t1")
            with c2_c:m_m_c_t1=st.number_input("💶 Montant min (€)",0.0,value=default_min_amt_t1,step=50.0,format="%.2f",key="m_m_c_t1")

            if st.button("🚀 Calculer Qtés Cmd",key="calc_q_c_b_t1"):
                curr_calc_params_t1 = {'suppliers': sel_f_t1, 'duree_semaines': d_s_c_t1, 'min_amount': m_m_c_t1, 'sem_cols_hash': hash(tuple(id_sem_cols))}
                st.session_state.commande_params_calculated_for = curr_calc_params_t1
                with st.spinner("Calcul qtés..."):res_c_t1=calculer_quantite_a_commander(df_disp_t1,id_sem_cols,m_m_c_t1,d_s_c_t1)
                if res_c_t1:
                    st.success("✅ Calcul qtés OK.");q_c_res,vN1_res,v12N1_res,v12l_res,m_c_res=res_c_t1
                    df_r_c_res=df_disp_t1.copy();df_r_c_res["Qte Cmdée"]=q_c_res
                    df_r_c_res["Vts N-1 Total (calc)"]=vN1_res;df_r_c_res["Vts 12 N-1 Sim (calc)"]=v12N1_res;df_r_c_res["Vts 12 Dern. (calc)"]=v12l_res
                    df_r_c_res["Tarif Ach."]=pd.to_numeric(df_r_c_res["Tarif d'achat"],errors='coerce').fillna(0)
                    df_r_c_res["Total Cmd (€)"]=df_r_c_res["Tarif Ach."]*df_r_c_res["Qte Cmdée"]
                    df_r_c_res["Stock Terme"]=df_r_c_res["Stock"]+df_r_c_res["Qte Cmdée"]
                    st.session_state.commande_result_df=df_r_c_res;st.session_state.commande_calculated_total_amount=m_c_res
                    st.session_state.commande_suppliers_calculated_for=sel_f_t1
                    st.rerun()
                else:st.error("❌ Calcul qtés échoué.")

            if st.session_state.commande_result_df is not None:
                curr_ui_params_t1_disp = {'suppliers': sel_f_t1, 'duree_semaines': d_s_c_t1, 'min_amount': m_m_c_t1, 'sem_cols_hash': hash(tuple(id_sem_cols))}
                if st.session_state.get('commande_params_calculated_for') == curr_ui_params_t1_disp:
                    st.markdown("---");st.markdown("#### Résultats Prévision Commande")
                    df_c_d_disp=st.session_state.commande_result_df;m_c_d_disp=st.session_state.commande_calculated_total_amount
                    st.metric(label="💰 Montant Total Cmd",value=f"{m_c_d_disp:,.2f} €")
                    if len(sel_f_t1)==1:
                        s_s_disp=sel_f_t1[0]
                        if s_s_disp in min_o_amts:
                            r_m_s_disp=min_o_amts[s_s_disp];a_t_s_disp=df_c_d_disp[df_c_d_disp["Fournisseur"]==s_s_disp]["Total Cmd (€)"].sum()
                            if r_m_s_disp>0 and a_t_s_disp<r_m_s_disp:st.warning(f"⚠️ Min non atteint ({s_s_disp}): {a_t_s_disp:,.2f}€ / Requis: {r_m_s_disp:,.2f}€ (Manque: {r_m_s_disp-a_t_s_disp:,.2f}€)")

                    cols_s_c_disp=["Fournisseur","AF_RefFourniss","Référence Article","Désignation Article","Stock","Vts N-1 Total (calc)","Vts 12 N-1 Sim (calc)","Vts 12 Dern. (calc)","Conditionnement","Qte Cmdée","Stock Terme","Tarif Ach.","Total Cmd (€)"]
                    disp_c_c_final=[c for c in cols_s_c_disp if c in df_c_d_disp.columns]
                    if not disp_c_c_final:st.error("Aucune col à afficher (cmd).")
                    else:
                        fmts_c_disp={"Tarif Ach.":"{:,.2f}€","Total Cmd (€)":"{:,.2f}€","Vts N-1 Total (calc)":"{:,.0f}","Vts 12 N-1 Sim (calc)":"{:,.0f}","Vts 12 Dern. (calc)":"{:,.0f}","Stock":"{:,.0f}","Conditionnement":"{:,.0f}","Qte Cmdée":"{:,.0f}","Stock Terme":"{:,.0f}"}
                        st.dataframe(df_c_d_disp[disp_c_c_final].style.format(fmts_c_disp,na_rep="-",thousands=","))

                    st.markdown("#### Export Commandes")
                    df_e_c_exp=df_c_d_disp[df_c_d_disp["Qte Cmdée"]>0].copy()
                    if not df_e_c_exp.empty:
                        out_b_c_exp=io.BytesIO();shts_c_exp=0
                        try:
                            with pd.ExcelWriter(out_b_c_exp,engine="openpyxl") as writer_c_exp:
                                exp_c_s_c_exp=[c for c in disp_c_c_final if c!='Fournisseur']
                                q_exp,p_exp,t_exp="Qte Cmdée","Tarif Ach.","Total Cmd (€)"
                                f_ok_exp=False
                                if all(c_exp in exp_c_s_c_exp for c_exp in[q_exp,p_exp,t_exp]):
                                    try:q_l_exp,p_l_exp,t_l_exp=get_column_letter(exp_c_s_c_exp.index(q_exp)+1),get_column_letter(exp_c_s_c_exp.index(p_exp)+1),get_column_letter(exp_c_s_c_exp.index(t_exp)+1);f_ok_exp=True
                                    except ValueError:pass
                                for sup_e_exp in sel_f_t1:
                                    df_s_e_exp=df_e_c_exp[df_e_c_exp["Fournisseur"]==sup_e_exp]
                                    if not df_s_e_exp.empty:
                                        df_w_s_exp=df_s_e_exp[exp_c_s_c_exp].copy();n_r_exp=len(df_w_s_exp);s_nm_exp=sanitize_sheet_name(sup_e_exp)
                                        df_w_s_exp.to_excel(writer_c_exp,sheet_name=s_nm_exp,index=False)
                                        ws_exp=writer_c_exp.sheets[s_nm_exp]
                                        cmd_col_fmts_exp={"Stock":"#,##0","Vts N-1 Total (calc)":"#,##0","Vts 12 N-1 Sim (calc)":"#,##0","Vts 12 Dern. (calc)":"#,##0","Conditionnement":"#,##0","Qte Cmdée":"#,##0","Stock Terme":"#,##0","Tarif Ach.":"#,##0.00€"}
                                        format_excel_sheet(ws_exp,df_w_s_exp,column_formats=cmd_col_fmts_exp)
                                        if f_ok_exp and n_r_exp>0:
                                            for r_idx_exp in range(2,n_r_exp+2):cell_t_exp=ws_exp[f"{t_l_exp}{r_idx_exp}"];cell_t_exp.value=f"={q_l_exp}{r_idx_exp}*{p_l_exp}{r_idx_exp}";cell_t_exp.number_format='#,##0.00€'
                                        lbl_name_col_exp="Désignation Article"
                                        if lbl_name_col_exp not in exp_c_s_c_exp: lbl_name_col_exp = exp_c_s_c_exp[1] if len(exp_c_s_c_exp)>1 else exp_c_s_c_exp[0]
                                        lbl_col_idx_excel = exp_c_s_c_exp.index(lbl_name_col_exp)+1
                                        total_col_idx_excel = exp_c_s_c_exp.index(t_exp)+1

                                        total_row_xl_idx_exp=n_r_exp+2
                                        ws_exp.cell(row=total_row_xl_idx_exp, column=lbl_col_idx_excel, value="TOTAL").font=Font(bold=True)
                                        cell_gt_exp=ws_exp.cell(row=total_row_xl_idx_exp, column=total_col_idx_excel)
                                        if n_r_exp>0:cell_gt_exp.value=f"=SUM({t_l_exp}2:{t_l_exp}{n_r_exp+1})"
                                        else:cell_gt_exp.value=0
                                        cell_gt_exp.number_format='#,##0.00€';cell_gt_exp.font=Font(bold=True)

                                        min_req_row_xl_idx_exp=n_r_exp+3
                                        ws_exp.cell(row=min_req_row_xl_idx_exp, column=lbl_col_idx_excel, value="Min Requis Fourn.").font=Font(bold=True)
                                        cell_min_req_v_exp=ws_exp.cell(row=min_req_row_xl_idx_exp, column=total_col_idx_excel)
                                        min_r_s_val_exp=min_o_amts.get(sup_e_exp,0);min_d_s_val_exp=f"{min_r_s_val_exp:,.2f}€"if min_r_s_val_exp>0 else"N/A"
                                        cell_min_req_v_exp.value=min_d_s_val_exp;cell_min_req_v_exp.font=Font(bold=True)

                                        if st.session_state.supplier_evaluation_data:
                                            supplier_eval_info_export = st.session_state.supplier_evaluation_data.get(sup_e_exp)
                                            if supplier_eval_info_export:
                                                target_stock_val_export = supplier_eval_info_export.get('max_stock_target', 0)
                                                target_stock_row_idx_excel = min_req_row_xl_idx_exp + 1
                                                ws_exp.cell(row=target_stock_row_idx_excel, column=lbl_col_idx_excel, value="Objectif Val. Stock Max Fourn.").font = Font(bold=True)
                                                cell_target_stock_val_excel = ws_exp.cell(row=target_stock_row_idx_excel, column=total_col_idx_excel)
                                                cell_target_stock_val_excel.value = f"{target_stock_val_export:,.2f}€"
                                                cell_target_stock_val_excel.font = Font(bold=True)
                                        shts_c_exp+=1
                            if shts_c_exp>0:
                                out_b_c_exp.seek(0)
                                fn_c_exp=f"commandes_{'multi'if len(sel_f_t1)>1 else sanitize_sheet_name(sel_f_t1[0])}_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx"
                                st.download_button(f"📥 Télécharger ({shts_c_exp} feuilles)",out_b_c_exp,fn_c_exp,"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",key="dl_c_b_t1_dl")
                            else:st.info("Aucune qté > 0 à exporter (ou err création feuilles).")
                        except Exception as e_wrt_c_exp:logging.exception(f"Err ExcelWriter cmd: {e_wrt_c_exp}");st.error("Erreur export commandes.")
                    else:st.info("Aucun article qté > 0 à exporter.")
                else:st.info("Paramètres changés. Relancer calcul pour résultats à jour.")

    # --- Tab 1 AI: Prévision Commande (IA) ---
    with tab1_ai:
        st.header("🤖 Prévision des Quantités à Commander (avec IA)")
        if not PROPHET_AVAILABLE:
            st.error("La librairie Prophet (pour l'IA) n'est pas installée. Cette fonctionnalité est désactivée.")
        else:
            sel_f_t1_ai = render_supplier_checkboxes("tab1_ai", all_sups_data, default_select_all=True)
            df_disp_t1_ai = pd.DataFrame()
            if sel_f_t1_ai:
                if not df_base_tabs.empty:
                    df_disp_t1_ai = df_base_tabs[df_base_tabs["Fournisseur"].isin(sel_f_t1_ai)].copy()
                    st.caption(f"{len(df_disp_t1_ai)} art. / {len(sel_f_t1_ai)} fourn.")

            if sel_f_t1_ai and not df_disp_t1_ai.empty:
                try:
                    stock_actuel_selection_ai = pd.to_numeric(df_disp_t1_ai["Stock"], errors='coerce').fillna(0)
                    tarif_achat_selection_ai = pd.to_numeric(df_disp_t1_ai["Tarif d'achat"], errors='coerce').fillna(0)
                    valeur_stock_selection_ai = (stock_actuel_selection_ai * tarif_achat_selection_ai).sum()
                    st.metric(label="📊 Valeur Stock Actuel (€) (Fourn. Sél.)", value=f"{valeur_stock_selection_ai:,.2f} €")
                except KeyError as e_stockval: st.error(f"Erreur : Colonne manquante pour valeur stock ('{e_stockval}').")
                except Exception as e_stockval_calc: st.error(f"Erreur calcul valeur stock actuel : {e_stockval_calc}")

                if len(sel_f_t1_ai) == 1 and st.session_state.supplier_evaluation_data:
                    supplier_name_selected = sel_f_t1_ai[0]
                    eval_data = st.session_state.supplier_evaluation_data.get(supplier_name_selected)
                    if eval_data:
                        st.metric(
                            label=f"🎯 Objectif Val. Stock Max pour {supplier_name_selected} (€)",
                            value=f"{eval_data.get('max_stock_target', 0):,.2f} €",
                            help=f"Basé sur {eval_data.get('cogs_pct', 0):.2f}% du CA d'Achat global ({eval_data.get('cogs_54w',0):,.0f}€ sur 54 sem.) et un objectif total de {st.session_state.global_stock_target_config:,.0f}€."
                        )
            elif sel_f_t1_ai and df_disp_t1_ai.empty:
                 st.metric(label="📊 Valeur Stock Actuel (€) (Fourn. Sél.)", value="0,00 €")
            else:
                st.info("Sélectionner fournisseur(s).")

            st.markdown("---")

            if df_disp_t1_ai.empty and sel_f_t1_ai:
                st.warning("Aucun article trouvé pour le(s) fournisseur(s) sélectionné(s).")
            elif not id_sem_cols and not df_disp_t1_ai.empty:
                st.warning("Colonnes ventes historiques non identifiées. Prévision IA impossible.")
            elif not df_disp_t1_ai.empty:
                st.markdown("#### Paramètres Prévision IA")
                c1_ai, c2_ai = st.columns(2)
                with c1_ai:
                    fcst_w_ai_t1 = st.number_input("⏳ Semaines à prévoir:", 1, 52, value=st.session_state.ai_forecast_weeks_val, step=1, key="fcst_w_ai_t1_numin")
                with c2_ai:
                    min_amt_ai_t1_default = st.session_state.ai_min_order_val
                    if len(sel_f_t1_ai) == 1 and sel_f_t1_ai[0] in min_o_amts and min_amt_ai_t1_default == 0.0:
                        min_amt_ai_t1_default = min_o_amts[sel_f_t1_ai[0]]
                    min_amt_ai_t1 = st.number_input("💶 Montant min (€) (si 1 fourn.):", 0.0, value=min_amt_ai_t1_default, step=50.0, format="%.2f", key="min_amt_ai_t1_numin")

                st.session_state.ai_forecast_weeks_val = fcst_w_ai_t1
                st.session_state.ai_min_order_val = min_amt_ai_t1

                st.markdown("##### Exclure des Fournisseurs de l'Ajustement Objectif Stock Max")
                suppliers_available_for_exclusion_t1_ai = []
                if sel_f_t1_ai and not df_disp_t1_ai.empty and 'Fournisseur' in df_disp_t1_ai.columns:
                    suppliers_available_for_exclusion_t1_ai = sorted(df_disp_t1_ai['Fournisseur'].astype(str).unique().tolist())
                
                if not suppliers_available_for_exclusion_t1_ai:
                    st.caption("Aucun fournisseur actuellement sélectionné pour la prévision IA ou données fournisseur manquantes.")
                else:
                    valid_excluded_suppliers = [
                        s for s in st.session_state.get('ai_excluded_suppliers_stock_target', []) 
                        if s in suppliers_available_for_exclusion_t1_ai
                    ]
                    if st.session_state.get('ai_excluded_suppliers_stock_target') != valid_excluded_suppliers:
                        st.session_state.ai_excluded_suppliers_stock_target = valid_excluded_suppliers
                    
                    excluded_suppliers_ui_t1_ai = st.multiselect(
                        "Fournisseurs à NE PAS soumettre à la règle de valeur de stock max:",
                        options=suppliers_available_for_exclusion_t1_ai,
                        default=st.session_state.ai_excluded_suppliers_stock_target,
                        key="excluded_supp_stock_target_ui_t1_ai",
                        help="Les commandes pour ces fournisseurs ne seront pas réduites si elles dépassent leur objectif de valeur de stock."
                    )
                    st.session_state.ai_excluded_suppliers_stock_target = excluded_suppliers_ui_t1_ai

                if st.button("🚀 Calculer Qtés avec IA", key="calc_q_ai_b_t1_go"):
                    events_hash = 0
                    df_events_state = st.session_state.get('df_product_events', pd.DataFrame())
                    if not df_events_state.empty:
                        try:
                            # S'assurer que les colonnes sont dans un ordre constant pour le hash
                            df_events_state_sorted = df_events_state.sort_index(axis=1)
                            events_hash = pd.util.hash_pandas_object(df_events_state_sorted, index=True).sum()
                        except Exception as e_hash:
                            logging.warning(f"Could not hash df_product_events: {e_hash}")


                    curr_calc_params_t1_ai = {
                        'suppliers': sel_f_t1_ai,
                        'forecast_weeks': fcst_w_ai_t1,
                        'min_amount_ui': min_amt_ai_t1,
                        'excluded_suppliers_stock_target': st.session_state.get('ai_excluded_suppliers_stock_target', []), 
                        'sem_cols_hash': hash(tuple(id_sem_cols)),
                        'events_hash': events_hash 
                    }
                    st.session_state.ai_commande_params_calculated_for = curr_calc_params_t1_ai

                    res_dfs_list_ai_calc = []
                    calc_ok_overall_ai = True
                    st.info(f"Lancement prévision IA pour {len(sel_f_t1_ai)} fournisseur(s)...")

                    for sup_idx_ai, sup_name_proc_ai in enumerate(sel_f_t1_ai):
                        df_sup_subset_ai_proc = df_disp_t1_ai[df_disp_t1_ai["Fournisseur"] == sup_name_proc_ai].copy()
                        sup_specific_min_order_ai = min_amt_ai_t1 if len(sel_f_t1_ai) == 1 else min_o_amts.get(sup_name_proc_ai, 0.0)
                        if not df_sup_subset_ai_proc.empty:
                            ai_res_df_sup, _ = ai_calculate_order_quantities(
                                df_sup_subset_ai_proc, 
                                id_sem_cols, 
                                fcst_w_ai_t1, 
                                sup_specific_min_order_ai,
                                apply_special_rules=True, # Garder la règle de stock bas / cond.
                                df_events_global=st.session_state.get('df_product_events', pd.DataFrame())
                            )
                            if ai_res_df_sup is not None: res_dfs_list_ai_calc.append(ai_res_df_sup)
                            else: st.error(f"Échec calcul IA pour: {sup_name_proc_ai}"); calc_ok_overall_ai = False
                        else: logging.info(f"Aucun article pour {sup_name_proc_ai} (IA).")

                    df_final_after_all_filters = pd.DataFrame() # Doit être défini avant le bloc if/else

                    if calc_ok_overall_ai and res_dfs_list_ai_calc:
                        final_ai_res_df_calc = pd.concat(res_dfs_list_ai_calc, ignore_index=True) if res_dfs_list_ai_calc else pd.DataFrame()
                        st.success("✅ Calcul IA initial terminé!")
                        df_before_350_filter = final_ai_res_df_calc.copy()
                        st.markdown("---")
                        st.info("Application du filtre : Commandes fournisseur < 350€ ignorées (sauf si article en stock < 0).")
                        df_after_350_filter = pd.DataFrame()
                        if not df_before_350_filter.empty:
                            for col_num_350 in ['Total Cmd (€) (IA)', 'Qté Cmdée (IA)', 'Stock']:
                                if col_num_350 in df_before_350_filter.columns:
                                     df_before_350_filter[col_num_350] = pd.to_numeric(df_before_350_filter[col_num_350], errors='coerce').fillna(0)
                            order_value_per_supplier = df_before_350_filter[df_before_350_filter['Qté Cmdée (IA)'] > 0].groupby('Fournisseur')['Total Cmd (€) (IA)'].sum()
                            suppliers_with_neg_stock_ordered = df_before_350_filter[(df_before_350_filter['Qté Cmdée (IA)'] > 0) & (df_before_350_filter['Stock'] < 0)]['Fournisseur'].unique()
                            suppliers_to_keep = set(s for s, v in order_value_per_supplier.items() if v >= 350 or s in suppliers_with_neg_stock_ordered)
                            initial_rows_350 = len(df_before_350_filter)
                            df_after_350_filter = df_before_350_filter[df_before_350_filter['Fournisseur'].isin(suppliers_to_keep)].copy()
                            filtered_rows_350 = len(df_after_350_filter)
                            if initial_rows_350 > filtered_rows_350: st.caption(f"{initial_rows_350 - filtered_rows_350} lignes article (< 350€ sans stock négatif) retirées.")
                            ignored_indices = df_before_350_filter.index.difference(df_after_350_filter.index)
                            df_ignored_orders_raw = df_before_350_filter.loc[ignored_indices].copy()
                            if not df_ignored_orders_raw.empty:
                                df_ignored_orders_raw['Qté Cmdée (IA)'] = pd.to_numeric(df_ignored_orders_raw['Qté Cmdée (IA)'], errors='coerce').fillna(0)
                                df_ignored_orders_filtered = df_ignored_orders_raw[df_ignored_orders_raw['Qté Cmdée (IA)'] > 0].copy()
                            else: df_ignored_orders_filtered = pd.DataFrame()
                            st.session_state.ai_ignored_orders_df = df_ignored_orders_filtered
                        else:
                             df_after_350_filter = df_before_350_filter.copy() # S'assurer que c'est une copie
                             st.session_state.ai_ignored_orders_df = pd.DataFrame()
                        
                        df_final_after_all_filters = df_after_350_filter.copy()

                        if st.session_state.supplier_evaluation_data and not df_final_after_all_filters.empty:
                            st.markdown("---")
                            st.info("Ajustement des commandes pour respecter les objectifs de valeur de stock max par fournisseur.")
                            suppliers_in_current_command = df_final_after_all_filters['Fournisseur'].unique()
                            df_all_items_for_selected_suppliers_ui = df_disp_t1_ai[df_disp_t1_ai['Fournisseur'].isin(suppliers_in_current_command)].copy()
                            df_to_adjust_iteratively = df_final_after_all_filters.copy()

                            excluded_suppliers_from_target_rule = st.session_state.get('ai_excluded_suppliers_stock_target', [])

                            for supplier_name_adj in suppliers_in_current_command:
                                if supplier_name_adj in excluded_suppliers_from_target_rule:
                                    st.caption(f"Fournisseur {supplier_name_adj} : Exclu de l'ajustement de l'objectif de stock max.")
                                    logging.info(f"Supplier {supplier_name_adj} is excluded from max stock target adjustment.")
                                    continue 

                                supplier_target_data = st.session_state.supplier_evaluation_data.get(supplier_name_adj)
                                if not supplier_target_data or 'max_stock_target' not in supplier_target_data:
                                    logging.warning(f"Pas de données d'objectif stock pour fournisseur {supplier_name_adj}."); continue
                                max_stock_target_for_supplier = supplier_target_data['max_stock_target']
                                df_supplier_all_items_current_disp_iter = df_all_items_for_selected_suppliers_ui[df_all_items_for_selected_suppliers_ui['Fournisseur'] == supplier_name_adj]
                                if df_supplier_all_items_current_disp_iter.empty: continue
                                current_stock_value_supplier = (pd.to_numeric(df_supplier_all_items_current_disp_iter['Stock'], errors='coerce').fillna(0) * pd.to_numeric(df_supplier_all_items_current_disp_iter["Tarif d'achat"], errors='coerce').fillna(0)).sum()
                                df_supplier_command_items_adj = df_to_adjust_iteratively[df_to_adjust_iteratively['Fournisseur'] == supplier_name_adj].copy()
                                if df_supplier_command_items_adj.empty: continue
                                for col_num in ['Stock', "Tarif d'achat", 'Qté Cmdée (IA)', 'Conditionnement']:
                                    if col_num in df_supplier_command_items_adj.columns: df_supplier_command_items_adj.loc[:, col_num] = pd.to_numeric(df_supplier_command_items_adj[col_num], errors='coerce').fillna(0)
                                df_supplier_command_items_adj.loc[:, 'Conditionnement'] = df_supplier_command_items_adj['Conditionnement'].apply(lambda x: int(x) if x > 0 else 1)
                                df_supplier_command_items_adj.loc[:, 'Qté Cmdée (IA)'] = df_supplier_command_items_adj['Qté Cmdée (IA)'].astype(int)
                                value_of_current_supplier_order = (df_supplier_command_items_adj['Qté Cmdée (IA)'] * df_supplier_command_items_adj["Tarif d'achat"]).sum()
                                projected_stock_value_supplier = current_stock_value_supplier + value_of_current_supplier_order
                                value_to_reduce_from_supplier_cmd = max(0, projected_stock_value_supplier - max_stock_target_for_supplier)
                                st.caption(f"Fourn: {supplier_name_adj} | Val.Stk Act: {current_stock_value_supplier:,.0f}€ | Val.Stk Proj (avant ajust.): {projected_stock_value_supplier:,.0f}€ | Cible Max: {max_stock_target_for_supplier:,.0f}€ | A Reduire Cmd: {value_to_reduce_from_supplier_cmd:,.0f}€")

                                if value_to_reduce_from_supplier_cmd > 0.01:
                                    wos_period_weeks = 12; available_weeks = len(id_sem_cols)
                                    weeks_to_use_for_wos_supplier = min(wos_period_weeks, available_weeks)
                                    df_supplier_command_items_adj['WoS_Calculated_Supplier'] = np.inf
                                    df_supplier_command_items_adj['SRM_Qty'] = 0
                                    if weeks_to_use_for_wos_supplier > 0:
                                        semaine_cols_for_wos_sup = id_sem_cols[-weeks_to_use_for_wos_supplier:]
                                        for item_idx_wos, item_row_wos in df_supplier_command_items_adj.iterrows():
                                            original_item_sales_data_series = pd.Series(dtype='float64')
                                            if item_idx_wos in df_base_tabs.index:
                                                sales_cols_present = [c for c in semaine_cols_for_wos_sup if c in df_base_tabs.columns]
                                                if sales_cols_present: original_item_sales_data_series = df_base_tabs.loc[item_idx_wos, sales_cols_present].fillna(0)
                                            else:
                                                ref_art_current_item = df_supplier_command_items_adj.loc[item_idx_wos, 'Référence Article']
                                                matching_row_in_base = df_base_tabs[df_base_tabs['Référence Article'] == ref_art_current_item]
                                                if not matching_row_in_base.empty:
                                                    sales_cols_present_fallback = [c for c in semaine_cols_for_wos_sup if c in matching_row_in_base.columns]
                                                    if sales_cols_present_fallback: original_item_sales_data_series = matching_row_in_base.iloc[0][sales_cols_present_fallback].fillna(0)
                                                else: logging.warning(f"Ventes non trouvées pour Art {item_idx_wos} / Ref {ref_art_current_item}")
                                            
                                            avg_weekly_sales_item = original_item_sales_data_series.sum() / weeks_to_use_for_wos_supplier if weeks_to_use_for_wos_supplier > 0 else 0
                                            current_stock_item_wos = item_row_wos['Stock']
                                            if avg_weekly_sales_item > 0: df_supplier_command_items_adj.loc[item_idx_wos, 'WoS_Calculated_Supplier'] = current_stock_item_wos / avg_weekly_sales_item
                                            elif current_stock_item_wos <= 0: df_supplier_command_items_adj.loc[item_idx_wos, 'WoS_Calculated_Supplier'] = 0.0
                                            srm_cond = item_row_wos['Conditionnement']
                                            srm_1wk_sales = np.ceil(avg_weekly_sales_item / srm_cond) * srm_cond if srm_cond > 0 else avg_weekly_sales_item
                                            df_supplier_command_items_adj.loc[item_idx_wos, 'SRM_Qty'] = max(srm_cond, srm_1wk_sales)
                                    
                                    candidates_reduc_supplier = df_supplier_command_items_adj[df_supplier_command_items_adj['Qté Cmdée (IA)'] > 0].copy()
                                    if not candidates_reduc_supplier.empty:
                                        candidates_reduc_supplier.sort_values(by='WoS_Calculated_Supplier', ascending=False, inplace=True, na_position='first')
                                        value_reduced_supplier_total = 0.0
                                        
                                        for item_index_reduc_sup in candidates_reduc_supplier.index:
                                            if value_to_reduce_from_supplier_cmd <= 0.01: break
                                            current_qty_reduc_sup = df_to_adjust_iteratively.loc[item_index_reduc_sup, 'Qté Cmdée (IA)']
                                            packaging_reduc_sup = df_to_adjust_iteratively.loc[item_index_reduc_sup, 'Conditionnement']
                                            price_reduc_sup = df_to_adjust_iteratively.loc[item_index_reduc_sup, "Tarif d'achat"]
                                            srm_sup = df_supplier_command_items_adj.loc[item_index_reduc_sup, 'SRM_Qty']
                                            if packaging_reduc_sup > 0 and price_reduc_sup > 0 and current_qty_reduc_sup > srm_sup :
                                                qty_reducible_above_srm_item = current_qty_reduc_sup - srm_sup
                                                num_pkgs_can_remove_item = int(qty_reducible_above_srm_item / packaging_reduc_sup)
                                                if num_pkgs_can_remove_item > 0:
                                                    value_per_pkg_reduc_sup = packaging_reduc_sup * price_reduc_sup
                                                    num_pkgs_to_reach_target_item = int(value_to_reduce_from_supplier_cmd / value_per_pkg_reduc_sup) if value_per_pkg_reduc_sup > 0 else 0
                                                    num_pkgs_actually_reduce_item = min(num_pkgs_can_remove_item, num_pkgs_to_reach_target_item)
                                                    if num_pkgs_actually_reduce_item == 0 and num_pkgs_can_remove_item > 0 and value_to_reduce_from_supplier_cmd > value_per_pkg_reduc_sup * 0.1:
                                                         num_pkgs_actually_reduce_item = 1
                                                    if num_pkgs_actually_reduce_item > 0:
                                                        qty_amount_to_reduce_sup = num_pkgs_actually_reduce_item * packaging_reduc_sup
                                                        value_of_this_reduction_sup = qty_amount_to_reduce_sup * price_reduc_sup
                                                        df_to_adjust_iteratively.loc[item_index_reduc_sup, 'Qté Cmdée (IA)'] -= qty_amount_to_reduce_sup
                                                        value_to_reduce_from_supplier_cmd -= value_of_this_reduction_sup
                                                        value_reduced_supplier_total += value_of_this_reduction_sup
                                        st.caption(f"Pour {supplier_name_adj}, réduction de {value_reduced_supplier_total:,.2f}€ appliquée (respectant SRM).")
                                        if value_to_reduce_from_supplier_cmd > 0.01: st.warning(f"Objectif stock pour {supplier_name_adj} non atteint. Excédent: {value_to_reduce_from_supplier_cmd:,.2f}€.")
                                    else: st.caption(f"Aucun article commandé/réductible pour {supplier_name_adj} pour son objectif stock.")
                                else: st.caption(f"Objectif de stock pour {supplier_name_adj} déjà respecté.")
                        
                        if not df_to_adjust_iteratively.empty:
                             df_to_adjust_iteratively['Total Cmd (€) (IA)'] = df_to_adjust_iteratively['Qté Cmdée (IA)'] * df_to_adjust_iteratively["Tarif d'achat"]
                             df_to_adjust_iteratively['Stock Terme (IA)'] = df_to_adjust_iteratively['Stock'] + df_to_adjust_iteratively['Qté Cmdée (IA)']
                        df_final_after_all_filters = df_to_adjust_iteratively.copy() # S'assurer que c'est une copie
                        
                        st.session_state.ai_commande_result_df = df_final_after_all_filters
                        st.session_state.ai_commande_total_amount = df_final_after_all_filters['Total Cmd (€) (IA)'].sum() if not df_final_after_all_filters.empty else 0.0
                        st.rerun()

                    elif not res_dfs_list_ai_calc: # Aucun résultat du tout
                        st.error("❌ Aucun résultat IA n'a pu être généré.")
                        st.session_state.ai_commande_result_df = pd.DataFrame(); st.session_state.ai_commande_total_amount = 0.0
                        st.session_state.ai_ignored_orders_df = pd.DataFrame()
                    else: # Calculs partiels échoués, mais certains résultats existent
                        st.warning("Certains calculs IA ont échoué. Filtre 350€ appliqué, ajustement objectif stock non appliqué sur résultats partiels.")
                        df_after_350_filter = pd.DataFrame(); df_ignored_partial = pd.DataFrame()
                        if res_dfs_list_ai_calc: # S'il y a au moins quelques résultats
                           final_ai_res_df_calc_partial = pd.concat(res_dfs_list_ai_calc, ignore_index=True)
                           df_before_350_partial = final_ai_res_df_calc_partial.copy()
                           if not df_before_350_partial.empty:
                               # Appliquer filtre 350€
                               for col_num_part in ['Total Cmd (€) (IA)', 'Qté Cmdée (IA)', 'Stock']:
                                   if col_num_part in df_before_350_partial.columns: df_before_350_partial[col_num_part] = pd.to_numeric(df_before_350_partial[col_num_part], errors='coerce').fillna(0)
                               order_value_per_supplier_part = df_before_350_partial[df_before_350_partial['Qté Cmdée (IA)'] > 0].groupby('Fournisseur')['Total Cmd (€) (IA)'].sum()
                               suppliers_with_neg_stock_ordered_part = df_before_350_partial[(df_before_350_partial['Qté Cmdée (IA)'] > 0) & (df_before_350_partial['Stock'] < 0)]['Fournisseur'].unique()
                               suppliers_to_keep_part = set(s for s, v in order_value_per_supplier_part.items() if v >= 350 or s in suppliers_with_neg_stock_ordered_part)
                               df_after_350_filter = df_before_350_partial[df_before_350_partial['Fournisseur'].isin(suppliers_to_keep_part)].copy()
                               
                               ignored_indices_partial = df_before_350_partial.index.difference(df_after_350_filter.index)
                               df_ignored_orders_raw_partial = df_before_350_partial.loc[ignored_indices_partial].copy()
                               if not df_ignored_orders_raw_partial.empty:
                                   df_ignored_orders_raw_partial['Qté Cmdée (IA)'] = pd.to_numeric(df_ignored_orders_raw_partial['Qté Cmdée (IA)'], errors='coerce').fillna(0)
                                   df_ignored_partial = df_ignored_orders_raw_partial[df_ignored_orders_raw_partial['Qté Cmdée (IA)'] > 0].copy()
                               else: df_ignored_partial = pd.DataFrame()
                           else: # df_before_350_partial était vide
                               df_after_350_filter = df_before_350_partial.copy() # df_after_350_filter sera aussi vide
                        
                        st.session_state.ai_commande_result_df = df_after_350_filter.copy()
                        st.session_state.ai_commande_total_amount = df_after_350_filter['Total Cmd (€) (IA)'].sum() if not df_after_350_filter.empty else 0.0
                        st.session_state.ai_ignored_orders_df = df_ignored_partial.copy()
                        st.rerun()

                if 'ai_commande_result_df' in st.session_state and st.session_state.ai_commande_result_df is not None:
                    events_hash_disp = 0
                    df_events_state_disp = st.session_state.get('df_product_events', pd.DataFrame())
                    if not df_events_state_disp.empty:
                        try:
                            df_events_state_disp_sorted = df_events_state_disp.sort_index(axis=1)
                            events_hash_disp = pd.util.hash_pandas_object(df_events_state_disp_sorted, index=True).sum()
                        except Exception as e_hash_disp:
                             logging.warning(f"Could not hash df_product_events for display check: {e_hash_disp}")
                    
                    curr_ui_params_t1_ai_disp = {
                        'suppliers': sel_f_t1_ai,
                        'forecast_weeks': fcst_w_ai_t1,
                        'min_amount_ui': min_amt_ai_t1,
                        'excluded_suppliers_stock_target': st.session_state.get('ai_excluded_suppliers_stock_target', []), 
                        'sem_cols_hash': hash(tuple(id_sem_cols)),
                        'events_hash': events_hash_disp 
                    }
                    if st.session_state.get('ai_commande_params_calculated_for') == curr_ui_params_t1_ai_disp:
                        st.markdown("---")
                        st.markdown("#### Résultats Prévision Commande (IA) - *Ajustés si nécessaire*")
                        df_disp_ai_res_final = st.session_state.ai_commande_result_df
                        total_amt_ai_res_final = st.session_state.ai_commande_total_amount

                        st.metric(label="💰 Montant Total Cmd (€) (IA)", value=f"{total_amt_ai_res_final:,.2f} €")

                        if not df_disp_ai_res_final.empty:
                            df_disp_ai_res_final['Stock'] = pd.to_numeric(df_disp_ai_res_final['Stock'], errors='coerce').fillna(0)
                            df_disp_ai_res_final['Qté Cmdée (IA)'] = pd.to_numeric(df_disp_ai_res_final['Qté Cmdée (IA)'], errors='coerce').fillna(0)
                            df_disp_ai_res_final["Tarif d'achat"] = pd.to_numeric(df_disp_ai_res_final["Tarif d'achat"], errors='coerce').fillna(0)
                            final_proj_stock_value = ((df_disp_ai_res_final['Stock'] + df_disp_ai_res_final['Qté Cmdée (IA)']) * df_disp_ai_res_final["Tarif d'achat"]).sum()
                            st.metric(label="📊 Valeur Stock Projeté (€) (Articles Commandés)", value=f"{final_proj_stock_value:,.2f} €")

                        for sup_chk_min_ai in sel_f_t1_ai:
                            sup_min_cfg_val_ai = min_o_amts.get(sup_chk_min_ai, 0.0)
                            min_applied_in_calc_ai = min_amt_ai_t1 if len(sel_f_t1_ai) == 1 else sup_min_cfg_val_ai
                            if min_applied_in_calc_ai > 0 and not df_disp_ai_res_final.empty:
                                actual_order_sup_ai = df_disp_ai_res_final[(df_disp_ai_res_final["Fournisseur"] == sup_chk_min_ai)]["Total Cmd (€) (IA)"].sum()
                                if actual_order_sup_ai < min_applied_in_calc_ai:
                                    st.warning(f"⚠️ Min cmd pour {sup_chk_min_ai} ({min_applied_in_calc_ai:,.2f}€) non atteint ({actual_order_sup_ai:,.2f}€) - *peut être dû à l'ajustement objectif stock*.")

                        cols_show_ai_res_final = ["Fournisseur","AF_RefFourniss","Référence Article","Désignation Article", "Stock", "Forecast Ventes (IA)"]
                        cols_show_ai_res_final.extend(["Conditionnement", "Qté Cmdée (IA)", "Stock Terme (IA)", "Tarif d'achat", "Total Cmd (€) (IA)"])
                        disp_cols_ai_final = [c for c in cols_show_ai_res_final if c in df_disp_ai_res_final.columns]

                        if not disp_cols_ai_final: st.error("Aucune col à afficher (résultats IA).")
                        else:
                            fmts_ai_final = {"Tarif d'achat":"{:,.2f}€","Total Cmd (€) (IA)":"{:,.2f}€","Forecast Ventes (IA)":"{:,.2f}","Stock":"{:,.0f}","Conditionnement":"{:,.0f}","Qté Cmdée (IA)":"{:,.0f}","Stock Terme (IA)":"{:,.0f}"}
                            df_display_ordered_only = df_disp_ai_res_final[df_disp_ai_res_final["Qté Cmdée (IA)"] > 0] if "Qté Cmdée (IA)" in df_disp_ai_res_final else df_disp_ai_res_final

                            if df_display_ordered_only.empty and not df_disp_ai_res_final.empty:
                                st.info("Aucune quantité à commander après application des filtres et objectifs.")
                            elif not df_display_ordered_only.empty :
                                st.dataframe(df_display_ordered_only[disp_cols_ai_final].style.format(fmts_ai_final,na_rep="-",thousands=","))
                            else:
                                st.dataframe(df_disp_ai_res_final[disp_cols_ai_final].style.format(fmts_ai_final,na_rep="-",thousands=","))

                        st.markdown("#### Export Commandes Prévision IA")
                        df_exp_ai_final_dl = df_disp_ai_res_final[df_disp_ai_res_final["Qté Cmdée (IA)"] > 0].copy()

                        if not df_exp_ai_final_dl.empty:
                            out_b_ai_exp_dl = io.BytesIO(); shts_ai_exp_dl = 0
                            try:
                                with pd.ExcelWriter(out_b_ai_exp_dl, engine="openpyxl") as writer_ai_exp_dl:
                                    exp_cols_sheet_ai_dl = [c for c in disp_cols_ai_final if c != 'Fournisseur']
                                    q_ai_dl, p_ai_dl, t_ai_dl = "Qté Cmdée (IA)", "Tarif d'achat", "Total Cmd (€) (IA)"
                                    f_ok_ai_dl = False
                                    if all(c_ai_dl in exp_cols_sheet_ai_dl for c_ai_dl in [q_ai_dl,p_ai_dl,t_ai_dl]):
                                        try: q_l_ai_dl,p_l_ai_dl,t_l_ai_dl=get_column_letter(exp_cols_sheet_ai_dl.index(q_ai_dl)+1),get_column_letter(exp_cols_sheet_ai_dl.index(p_ai_dl)+1),get_column_letter(exp_cols_sheet_ai_dl.index(t_ai_dl)+1);f_ok_ai_dl=True
                                        except ValueError: pass

                                    suppliers_in_final_export = df_exp_ai_final_dl['Fournisseur'].unique()
                                    for sup_e_ai_dl in suppliers_in_final_export:
                                        df_s_e_ai_dl=df_exp_ai_final_dl[df_exp_ai_final_dl["Fournisseur"]==sup_e_ai_dl]
                                        df_w_s_ai_dl=df_s_e_ai_dl[exp_cols_sheet_ai_dl].copy()
                                        n_r_ai_dl=len(df_w_s_ai_dl);s_nm_ai_dl=sanitize_sheet_name(f"IA_Cmd_{sup_e_ai_dl}")
                                        df_w_s_ai_dl.to_excel(writer_ai_exp_dl,sheet_name=s_nm_ai_dl,index=False)
                                        ws_ai_dl=writer_ai_exp_dl.sheets[s_nm_ai_dl]
                                        cmd_col_fmts_ai_dl={"Stock":"#,##0","Forecast Ventes (IA)":"#,##0.00","Conditionnement":"#,##0","Qté Cmdée (IA)":"#,##0","Stock Terme (IA)":"#,##0","Tarif d'achat":"#,##0.00€"}
                                        format_excel_sheet(ws_ai_dl,df_w_s_ai_dl,column_formats=cmd_col_fmts_ai_dl)
                                        if f_ok_ai_dl and n_r_ai_dl>0:
                                            for r_idx_ai_dl in range(2,n_r_ai_dl+2):cell_t_ai_dl=ws_ai_dl[f"{t_l_ai_dl}{r_idx_ai_dl}"];cell_t_ai_dl.value=f"={q_l_ai_dl}{r_idx_ai_dl}*{p_l_ai_dl}{r_idx_ai_dl}";cell_t_ai_dl.number_format='#,##0.00€'
                                        lbl_name_col_ai_dl="Désignation Article"
                                        if lbl_name_col_ai_dl not in exp_cols_sheet_ai_dl: lbl_name_col_ai_dl = exp_cols_sheet_ai_dl[1] if len(exp_cols_sheet_ai_dl)>1 else exp_cols_sheet_ai_dl[0]
                                        lbl_col_idx_excel_ai = exp_cols_sheet_ai_dl.index(lbl_name_col_ai_dl)+1
                                        total_col_idx_excel_ai = exp_cols_sheet_ai_dl.index(t_ai_dl)+1
                                        total_row_xl_idx_ai_dl=n_r_ai_dl+2
                                        ws_ai_dl.cell(row=total_row_xl_idx_ai_dl, column=lbl_col_idx_excel_ai, value="TOTAL").font=Font(bold=True)
                                        cell_gt_ai_dl=ws_ai_dl.cell(row=total_row_xl_idx_ai_dl, column=total_col_idx_excel_ai)
                                        if n_r_ai_dl>0:cell_gt_ai_dl.value=f"=SUM({t_l_ai_dl}2:{t_l_ai_dl}{n_r_ai_dl+1})"
                                        else:cell_gt_ai_dl.value=0
                                        cell_gt_ai_dl.number_format='#,##0.00€';cell_gt_ai_dl.font=Font(bold=True)
                                        min_req_row_xl_idx_ai_dl=n_r_ai_dl+3
                                        ws_ai_dl.cell(row=min_req_row_xl_idx_ai_dl, column=lbl_col_idx_excel_ai, value="Min Requis Fourn.").font=Font(bold=True)
                                        cell_min_req_v_ai_dl=ws_ai_dl.cell(row=min_req_row_xl_idx_ai_dl, column=total_col_idx_excel_ai)
                                        min_r_s_val_ai_dl=min_o_amts.get(sup_e_ai_dl,0);min_d_s_val_ai_dl=f"{min_r_s_val_ai_dl:,.2f}€"if min_r_s_val_ai_dl>0 else"N/A"
                                        cell_min_req_v_ai_dl.value=min_d_s_val_ai_dl;cell_min_req_v_ai_dl.font=Font(bold=True)
                                        if st.session_state.supplier_evaluation_data:
                                            supplier_eval_info_export_ai = st.session_state.supplier_evaluation_data.get(sup_e_ai_dl)
                                            if supplier_eval_info_export_ai:
                                                target_stock_val_export_ai = supplier_eval_info_export_ai.get('max_stock_target', 0)
                                                target_stock_row_idx_excel_ai = min_req_row_xl_idx_ai_dl + 1
                                                ws_ai_dl.cell(row=target_stock_row_idx_excel_ai, column=lbl_col_idx_excel_ai, value="Objectif Val. Stock Max Fourn.").font = Font(bold=True)
                                                cell_target_stock_val_excel_ai = ws_ai_dl.cell(row=target_stock_row_idx_excel_ai, column=total_col_idx_excel_ai)
                                                cell_target_stock_val_excel_ai.value = f"{target_stock_val_export_ai:,.2f}€"
                                                cell_target_stock_val_excel_ai.font = Font(bold=True)
                                        shts_ai_exp_dl+=1
                                if shts_ai_exp_dl > 0:
                                    out_b_ai_exp_dl.seek(0)
                                    fn_ai_dl=f"commandes_IA_validees_{'multi'if len(sel_f_t1_ai)>1 else sanitize_sheet_name(sel_f_t1_ai[0])}_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx"
                                    st.download_button(f"📥 Télécharger Commandes Validées ({shts_ai_exp_dl} feuilles)",out_b_ai_exp_dl,fn_ai_dl,"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",key="dl_ai_cmd_final_b_t1_dl")
                                else:st.info("Aucune qté IA > 0 à exporter après filtres.")
                            except Exception as e_wrt_ai_dl:logging.exception(f"Err ExcelWriter cmd IA: {e_wrt_ai_dl}");st.error("Erreur export commandes IA.")
                        else:st.info("Aucun article qté IA > 0 à exporter après filtres.")

                        if 'ai_ignored_orders_df' in st.session_state and st.session_state.ai_ignored_orders_df is not None and not st.session_state.ai_ignored_orders_df.empty:
                            st.markdown("---")
                            st.markdown("##### Export Commandes Ignorées par IA (< 350€ sans stock nég.)")
                            df_ignored_export = st.session_state.ai_ignored_orders_df
                            cols_ignored_export_base = ["AF_RefFourniss", "Référence Article", "Désignation Article", "Stock", "Forecast Ventes (IA)", "Conditionnement", "Qté Cmdée (IA)", "Total Cmd (€) (IA)"]
                            out_b_ignored = io.BytesIO(); sheets_ignored_count = 0
                            try:
                                with pd.ExcelWriter(out_b_ignored, engine="openpyxl") as writer_ignored:
                                    suppliers_in_ignored_export = df_ignored_export['Fournisseur'].unique()
                                    for sup_ignored in suppliers_in_ignored_export:
                                        df_sup_ignored = df_ignored_export[df_ignored_export["Fournisseur"] == sup_ignored]
                                        cols_to_write_sheet_ignored = [c for c in cols_ignored_export_base if c in df_sup_ignored.columns]
                                        df_sheet_ignored_export = df_sup_ignored[cols_to_write_sheet_ignored].copy()
                                        if not df_sheet_ignored_export.empty:
                                            for col_num_ign_exp in ["Stock", "Forecast Ventes (IA)", "Conditionnement", "Qté Cmdée (IA)", "Total Cmd (€) (IA)"]:
                                                if col_num_ign_exp in df_sheet_ignored_export.columns:
                                                    df_sheet_ignored_export[col_num_ign_exp] = pd.to_numeric(df_sheet_ignored_export[col_num_ign_exp], errors='coerce').fillna(0)
                                            sheet_name_ignored = sanitize_sheet_name(f"Ign_{sup_ignored}")
                                            df_sheet_ignored_export.to_excel(writer_ignored, sheet_name=sheet_name_ignored, index=False)
                                            ws_ignored = writer_ignored.sheets[sheet_name_ignored]
                                            ignored_fmts_excel = {"Stock":"#,##0", "Forecast Ventes (IA)":"#,##0.00", "Conditionnement":"#,##0", "Qté Cmdée (IA)":"#,##0", "Total Cmd (€) (IA)":"#,##0.00€"}
                                            format_excel_sheet(ws_ignored, df_sheet_ignored_export, column_formats=ignored_fmts_excel)
                                            n_r_ign = len(df_sheet_ignored_export)
                                            if n_r_ign > 0 and "Total Cmd (€) (IA)" in cols_to_write_sheet_ignored and "Désignation Article" in cols_to_write_sheet_ignored:
                                                t_col_ign_letter = get_column_letter(cols_to_write_sheet_ignored.index("Total Cmd (€) (IA)") + 1)
                                                lbl_col_ign_name = "Désignation Article" if "Désignation Article" in cols_to_write_sheet_ignored else (cols_to_write_sheet_ignored[0] if cols_to_write_sheet_ignored else "A")
                                                lbl_col_ign_idx = cols_to_write_sheet_ignored.index(lbl_col_ign_name) +1 if lbl_col_ign_name in cols_to_write_sheet_ignored else 1

                                                ws_ignored.cell(row=n_r_ign + 2, column=lbl_col_ign_idx, value="TOTAL IGNORÉ").font = Font(bold=True)
                                                cell_gt_ign = ws_ignored[f"{t_col_ign_letter}{n_r_ign + 2}"]
                                                cell_gt_ign.value = f"=SUM({t_col_ign_letter}2:{t_col_ign_letter}{n_r_ign + 1})"
                                                cell_gt_ign.number_format = '#,##0.00€'; cell_gt_ign.font = Font(bold=True)
                                            sheets_ignored_count += 1
                                if sheets_ignored_count > 0:
                                    out_b_ignored.seek(0)
                                    fn_ignored = f"commandes_IA_ignorees_par_fourn_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx"
                                    st.download_button(
                                        label=f"📥 Télécharger Commandes Ignorées ({sheets_ignored_count} feuille(s) fournisseur)",
                                        data=out_b_ignored, file_name=fn_ignored,
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                        key="dl_ai_cmd_ignored_multi_b_t1_dl"
                                    )
                            except Exception as e_wrt_ignored: logging.exception(f"Err ExcelWriter cmd Ignorées: {e_wrt_ignored}"); st.error("Erreur export commandes ignorées.")
                    else:st.info("Paramètres IA changés. Relancer calcul pour résultats à jour.")

    # --- Tab 2: Stock Rotation Analysis ---
    with tab2:
        # ... (Code identique à la version précédente) ...
        st.header("Analyse de la Rotation des Stocks")
        sel_f_t2 = render_supplier_checkboxes("tab2", all_sups_data, default_select_all=True)
        df_disp_t2 = pd.DataFrame()
        if sel_f_t2:
            if not df_base_tabs.empty: df_disp_t2 = df_base_tabs[df_base_tabs["Fournisseur"].isin(sel_f_t2)].copy(); st.caption(f"{len(df_disp_t2)} art. / {len(sel_f_t2)} fourn.")
        else: st.info("Sélectionner fournisseur(s).")
        st.markdown("---")
        if df_disp_t2.empty and sel_f_t2: st.warning("Aucun article pour fournisseur(s) sélectionné(s).")
        elif not id_sem_cols and not df_disp_t2.empty: st.warning("Colonnes ventes non identifiées.")
        elif not df_disp_t2.empty:
            st.markdown("#### Paramètres Analyse Rotation")
            c1_r_t2,c2_r_t2=st.columns(2);
            with c1_r_t2:
                p_opts_r_t2={"12 dernières semaines":12,"52 dernières semaines":52,"Total disponible":0}
                default_period_label_t2 = st.session_state.get('rotation_analysis_period_label', "12 dernières semaines")
                if default_period_label_t2 not in p_opts_r_t2: default_period_label_t2 = "12 dernières semaines"
                sel_p_lbl_r_t2=st.selectbox("⏳ Période analyse:",list(p_opts_r_t2.keys()), index=list(p_opts_r_t2.keys()).index(default_period_label_t2), key="r_p_sel_ui_t2")
                sel_p_w_r_t2=p_opts_r_t2[sel_p_lbl_r_t2]
            with c2_r_t2:
                st.markdown("##### Options Affichage")
                show_all_r_t2=st.checkbox("Afficher tout",value=st.session_state.show_all_rotation_data,key="show_all_r_ui_cb_t2")
                r_thr_ui_t2=st.number_input("... ou vts mens. <",0.0,value=st.session_state.rotation_threshold_value,step=0.1,format="%.1f",key="r_thr_ui_numin_t2",disabled=show_all_r_t2)
            st.session_state.rotation_analysis_period_label = sel_p_lbl_r_t2
            st.session_state.show_all_rotation_data = show_all_r_t2
            if not show_all_r_t2: st.session_state.rotation_threshold_value = r_thr_ui_t2
            if st.button("🔄 Analyser Rotation",key="analyze_r_btn_t2"):
                curr_calc_params_t2 = {'suppliers': sel_f_t2, 'period_label': sel_p_lbl_r_t2, 'show_all': show_all_r_t2, 'threshold': r_thr_ui_t2 if not show_all_r_t2 else -1, 'sem_cols_hash': hash(tuple(id_sem_cols))}
                st.session_state.rotation_params_calculated_for = curr_calc_params_t2
                with st.spinner("Analyse rotation..."):df_r_res_t2=calculer_rotation_stock(df_disp_t2,id_sem_cols,sel_p_w_r_t2)
                if df_r_res_t2 is not None: st.success("✅ Analyse rotation OK.");st.session_state.rotation_result_df=df_r_res_t2; st.rerun()
                else:st.error("❌ Analyse rotation échouée.")
            if st.session_state.rotation_result_df is not None:
                curr_ui_params_t2_disp = {'suppliers': sel_f_t2, 'period_label': sel_p_lbl_r_t2, 'show_all': show_all_r_t2, 'threshold': r_thr_ui_t2 if not show_all_r_t2 else -1, 'sem_cols_hash': hash(tuple(id_sem_cols))}
                if st.session_state.get('rotation_params_calculated_for') == curr_ui_params_t2_disp:
                    st.markdown("---");st.markdown(f"#### Résultats Rotation ({sel_p_lbl_r_t2})")
                    df_r_orig_t2=st.session_state.rotation_result_df
                    df_r_disp_t2_final=pd.DataFrame();df_r_to_fmt_t2_final=pd.DataFrame()
                    if df_r_orig_t2.empty:st.info("Aucune donnée rotation à afficher.")
                    elif show_all_r_t2: df_r_disp_t2_final=df_r_orig_t2.copy();df_r_to_fmt_t2_final=df_r_disp_t2_final.copy();st.caption(f"Affichage {len(df_r_disp_t2_final)} articles.")
                    else:
                        m_sales_c_r_t2="Ventes Moy Mensuel (Période)"
                        if m_sales_c_r_t2 in df_r_orig_t2.columns:
                            try:
                                sales_f_t2=pd.to_numeric(df_r_orig_t2[m_sales_c_r_t2],errors='coerce').fillna(0)
                                df_r_disp_t2_final=df_r_orig_t2[sales_f_t2 < r_thr_ui_t2].copy();df_r_to_fmt_t2_final=df_r_disp_t2_final.copy()
                                st.caption(f"Filtre: Vts < {r_thr_ui_t2:.1f}/mois. {len(df_r_disp_t2_final)} / {len(df_r_orig_t2)} art.")
                                if df_r_disp_t2_final.empty:st.info(f"Aucun article < {r_thr_ui_t2:.1f} vts/mois.")
                            except Exception as ef_r_t2:st.error(f"Err filtre: {ef_r_t2}");df_r_disp_t2_final=df_r_orig_t2.copy();df_r_to_fmt_t2_final=df_r_disp_t2_final.copy()
                        else:st.warning(f"Col '{m_sales_c_r_t2}' non trouvée. Affichage tout.");df_r_disp_t2_final=df_r_orig_t2.copy();df_r_to_fmt_t2_final=df_r_disp_t2_final.copy()
                    if not df_r_disp_t2_final.empty:
                        cols_r_s_t2=["Fournisseur","AF_RefFourniss","Référence Article","Désignation Article","Tarif d'achat","Stock","Unités Vendues (Période)","Ventes Moy Hebdo (Période)","Ventes Moy Mensuel (Période)","Semaines Stock (WoS)","Rotation Unités (Proxy)","Valeur Stock Actuel (€)","COGS (Période)","Rotation Valeur (Proxy)"]
                        disp_c_r_t2=[c for c in cols_r_s_t2 if c in df_r_disp_t2_final.columns]
                        df_d_cp_r_t2=df_r_disp_t2_final[disp_c_r_t2].copy()
                        num_rnd_r_t2={"Tarif d'achat":2,"Ventes Moy Hebdo (Période)":2,"Ventes Moy Mensuel (Période)":2,"Semaines Stock (WoS)":1,"Rotation Unités (Proxy)":2,"Valeur Stock Actuel (€)":2,"COGS (Période)":2,"Rotation Valeur (Proxy)":2}
                        for c_t2,d_t2 in num_rnd_r_t2.items():
                            if c_t2 in df_d_cp_r_t2.columns:df_d_cp_r_t2[c_t2]=pd.to_numeric(df_d_cp_r_t2[c_t2],errors='coerce').round(d_t2)
                        df_d_cp_r_t2.replace([np.inf,-np.inf],'Infini',inplace=True)
                        fmts_r_t2={"Tarif d'achat":"{:,.2f}€","Stock":"{:,.0f}","Unités Vendues (Période)":"{:,.0f}","Ventes Moy Hebdo (Période)":"{:,.2f}","Ventes Moy Mensuel (Période)":"{:,.2f}","Semaines Stock (WoS)":"{}","Rotation Unités (Proxy)":"{}","Valeur Stock Actuel (€)":"{:,.2f}€","COGS (Période)":"{:,.2f}€","Rotation Valeur (Proxy)":"{}"}
                        st.dataframe(df_d_cp_r_t2.style.format(fmts_r_t2,na_rep="-",thousands=","))
                        st.markdown("#### Export Analyse Affichée")
                        if not df_r_to_fmt_t2_final.empty:
                            out_b_r_t2_exp=io.BytesIO();df_e_r_t2_exp=df_r_to_fmt_t2_final[disp_c_r_t2].copy()
                            df_e_r_t2_exp.replace([np.inf, -np.inf], "Infini", inplace=True)
                            lbl_e_r_t2=f"Filtree_{r_thr_ui_t2:.1f}"if not show_all_r_t2 else"Complete";sh_nm_r_t2=sanitize_sheet_name(f"Rotation_{lbl_e_r_t2}");f_base_r_t2=f"analyse_rotation_{lbl_e_r_t2}"
                            sup_e_nm_r_t2='multi'if len(sel_f_t2)>1 else(sanitize_sheet_name(sel_f_t2[0])if sel_f_t2 else'NA')
                            try:
                                with pd.ExcelWriter(out_b_r_t2_exp,engine="openpyxl")as wr_r_t2:
                                    df_e_r_t2_exp.to_excel(wr_r_t2,sheet_name=sh_nm_r_t2,index=False)
                                    ws_r_t2=wr_r_t2.sheets[sh_nm_r_t2]
                                    rot_col_fmts_t2={"Tarif d'achat":"#,##0.00€","Stock":"#,##0","Unités Vendues (Période)":"#,##0","Ventes Moy Hebdo (Période)":"#,##0.00","Ventes Moy Mensuel (Période)":"#,##0.00","Semaines Stock (WoS)":"0.0","Rotation Unités (Proxy)":"0.00","Valeur Stock Actuel (€)":"#,##0.00€","COGS (Période)":"#,##0.00€","Rotation Valeur (Proxy)":"0.00"}
                                    format_excel_sheet(ws_r_t2,df_e_r_t2_exp,column_formats=rot_col_fmts_t2)
                                out_b_r_t2_exp.seek(0);f_r_exp_t2=f"{f_base_r_t2}_{sup_e_nm_r_t2}_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx"
                                dl_lbl_r_t2=f"📥 Télécharger ({'Filtrée'if not show_all_r_t2 else'Complète'})"
                                st.download_button(dl_lbl_r_t2,out_b_r_t2_exp,f_r_exp_t2,"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",key="dl_r_b_t2_dl")
                            except Exception as e_wrt_r_t2:logging.exception(f"Err ExcelWriter rot: {e_wrt_r_t2}");st.error("Erreur export rotation.")
                        else:st.info("Aucune donnée à exporter.")
                else:st.info("Paramètres analyse rotation changés. Relancer analyse.")

    # --- Tab 3: Negative Stock Check ---
    with tab3:
        # ... (Code identique à la version précédente) ...
        st.header("Vérification des Stocks Négatifs")
        st.caption("Analyse tous articles du 'Tableau final'.")
        df_full_neg_t3=st.session_state.get('df_full',None)
        if df_full_neg_t3 is None or not isinstance(df_full_neg_t3,pd.DataFrame):st.warning("Données non chargées.")
        elif df_full_neg_t3.empty:st.info("'Tableau final' vide.")
        else:
            stock_c_neg_t3="Stock"
            if stock_c_neg_t3 not in df_full_neg_t3.columns:st.error(f"Colonne '{stock_c_neg_t3}' non trouvée.")
            else:
                df_neg_res_t3=df_full_neg_t3[pd.to_numeric(df_full_neg_t3[stock_c_neg_t3], errors='coerce').fillna(0)<0].copy()
                if df_neg_res_t3.empty:st.success("✅ Aucun stock négatif.")
                else:
                    st.warning(f"⚠️ **{len(df_neg_res_t3)} article(s) avec stock négatif !**")
                    cols_neg_show_t3=["Fournisseur","AF_RefFourniss","Référence Article","Désignation Article",stock_c_neg_t3]
                    disp_cols_neg_t3=[c for c in cols_neg_show_t3 if c in df_neg_res_t3.columns]
                    if not disp_cols_neg_t3:st.error("Cols manquantes affichage négatifs.")
                    else:
                        def highlight_negative(s):
                            is_negative = pd.to_numeric(s, errors='coerce') < 0
                            return ['background-color: #FADBD8' if v else '' for v in is_negative]
                        st.dataframe(df_neg_res_t3[disp_cols_neg_t3].style.format({stock_c_neg_t3:"{:,.0f}"},na_rep="-").apply(highlight_negative, subset=[stock_c_neg_t3], axis=0))
                    st.markdown("---");st.markdown("#### Exporter Stocks Négatifs")
                    out_b_neg_t3=io.BytesIO();df_exp_neg_t3=df_neg_res_t3[disp_cols_neg_t3].copy()
                    try:
                        with pd.ExcelWriter(out_b_neg_t3,engine="openpyxl")as w_neg_t3:
                            df_exp_neg_t3.to_excel(w_neg_t3,sheet_name="Stocks_Negatifs",index=False)
                            ws_neg_t3=w_neg_t3.sheets["Stocks_Negatifs"]
                            neg_col_fmts_t3={stock_c_neg_t3:"#,##0"}
                            format_excel_sheet(ws_neg_t3,df_exp_neg_t3,column_formats=neg_col_fmts_t3)
                        out_b_neg_t3.seek(0);f_neg_exp_t3=f"stocks_negatifs_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx"
                        st.download_button("📥 Télécharger Liste Négatifs",out_b_neg_t3,f_neg_exp_t3,"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",key="dl_neg_b_t3_dl")
                    except Exception as e_exp_neg_t3:st.error(f"Err export neg: {e_exp_neg_t3}")

    # --- Tab 4: Forecast Simulation ---
    with tab4:
        # ... (Code identique à la version précédente) ...
        st.header("Simulation de Forecast Annuel")
        sel_f_t4 = render_supplier_checkboxes("tab4", all_sups_data, default_select_all=True)
        df_disp_t4 = pd.DataFrame()
        if sel_f_t4:
            if not df_base_tabs.empty: df_disp_t4 = df_base_tabs[df_base_tabs["Fournisseur"].isin(sel_f_t4)].copy(); st.caption(f"{len(df_disp_t4)} art. / {len(sel_f_t4)} fourn.")
        else: st.info("Sélectionner fournisseur(s).")
        st.markdown("---");st.warning("🚨 **Hypothèse:** Saisonnalité mensuelle approx. sur 52 sem. N-1.")
        if df_disp_t4.empty and sel_f_t4: st.warning("Aucun article pour fournisseur(s) sélectionné(s).")
        elif len(id_sem_cols)<52 and not df_disp_t4.empty: st.warning(f"Données histo. < 52 sem ({len(id_sem_cols)}). Simu N-1 impossible.")
        elif not df_disp_t4.empty:
            st.markdown("#### Paramètres Simulation Forecast")
            all_cal_m_t4=list(calendar.month_name)[1:]
            sel_m_f_ui_t4=st.multiselect("📅 Mois simulation:",all_cal_m_t4,default=st.session_state.forecast_selected_months_ui,key="f_m_sel_ui_t4")
            sim_t_opts_f_t4=('Simple Progression','Objectif Montant')
            current_sim_type_index_t4 = st.session_state.get('forecast_sim_type_radio_index', 0)
            sim_t_f_ui_t4=st.radio("⚙️ Type Simulation:",sim_t_opts_f_t4,horizontal=True,index=current_sim_type_index_t4,key="f_sim_t_ui_t4")
            prog_pct_f_t4,obj_mt_f_t4=0.0,0.0
            c1_f_t4,c2_f_t4=st.columns(2);
            with c1_f_t4:
                if sim_t_f_ui_t4=='Simple Progression':
                    prog_pct_f_t4=st.number_input("📈 Progression (%)",-100.0,value=st.session_state.forecast_progression_percentage_ui,step=0.5,format="%.1f",key="f_prog_pct_ui_t4")
            with c2_f_t4:
                if sim_t_f_ui_t4=='Objectif Montant':
                    obj_mt_f_t4=st.number_input("🎯 Objectif (€) (mois sel.)",0.0,value=st.session_state.forecast_target_amount_ui,step=1000.0,format="%.2f",key="f_target_amt_ui_t4")
            st.session_state.forecast_selected_months_ui = sel_m_f_ui_t4
            st.session_state.forecast_sim_type_radio_index = sim_t_opts_f_t4.index(sim_t_f_ui_t4)
            if sim_t_f_ui_t4=='Simple Progression': st.session_state.forecast_progression_percentage_ui = prog_pct_f_t4
            if sim_t_f_ui_t4=='Objectif Montant': st.session_state.forecast_target_amount_ui = obj_mt_f_t4
            if st.button("▶️ Lancer Simulation Forecast",key="run_f_sim_btn_t4"):
                if not sel_m_f_ui_t4:st.error("Sélectionner au moins un mois.")
                else:
                    curr_calc_params_t4 = {'suppliers': sel_f_t4, 'months': sel_m_f_ui_t4, 'type': sim_t_f_ui_t4, 'prog_pct': prog_pct_f_t4, 'obj_amt': obj_mt_f_t4, 'sem_cols_hash': hash(tuple(id_sem_cols))}
                    st.session_state.forecast_simulation_params_calculated_for = curr_calc_params_t4
                    with st.spinner("Simulation forecast..."):df_f_res_t4,gt_f_t4=calculer_forecast_simulation_v3(df_disp_t4,id_sem_cols,sel_m_f_ui_t4,sim_t_f_ui_t4,prog_pct_f_t4,obj_mt_f_t4)
                    if df_f_res_t4 is not None: st.success("✅ Simu forecast OK.");st.session_state.forecast_result_df=df_f_res_t4;st.session_state.forecast_grand_total_amount=gt_f_t4; st.rerun()
                    else:st.error("❌ Simu forecast échouée.")
            if st.session_state.forecast_result_df is not None:
                curr_ui_params_t4_disp = {'suppliers': sel_f_t4, 'months': sel_m_f_ui_t4, 'type': sim_t_f_ui_t4, 'prog_pct': prog_pct_f_t4, 'obj_amt': obj_mt_f_t4, 'sem_cols_hash': hash(tuple(id_sem_cols))}
                if st.session_state.get('forecast_simulation_params_calculated_for') == curr_ui_params_t4_disp:
                    st.markdown("---");st.markdown("#### Résultats Simulation Forecast")
                    df_f_disp_t4=st.session_state.forecast_result_df;gt_f_disp_t4=st.session_state.forecast_grand_total_amount
                    if df_f_disp_t4.empty:st.info("Aucun résultat simulation.")
                    else:
                        fmts_f_t4={"Tarif d'achat":"{:,.2f}€","Conditionnement":"{:,.0f}"}
                        for m_disp_t4 in sel_m_f_ui_t4:
                            if f"Ventes N-1 {m_disp_t4}"in df_f_disp_t4.columns:fmts_f_t4[f"Ventes N-1 {m_disp_t4}"]="{:,.0f}"
                            if f"Qté Prév. {m_disp_t4}"in df_f_disp_t4.columns:fmts_f_t4[f"Qté Prév. {m_disp_t4}"]="{:,.0f}"
                            if f"Montant Prév. {m_disp_t4} (€)"in df_f_disp_t4.columns:fmts_f_t4[f"Montant Prév. {m_disp_t4} (€)"]="{:,.2f}€"
                        for col_n_t4 in["Vts N-1 Tot (Mois Sel.)","Qté Tot Prév (Mois Sel.)","Mnt Tot Prév (€) (Mois Sel.)"]:
                            if col_n_t4 in df_f_disp_t4.columns:fmts_f_t4[col_n_t4]="{:,.0f}"if"Qté"in col_n_t4 or"Vts"in col_n_t4 else"{:,.2f}€"
                        try:st.dataframe(df_f_disp_t4.style.format(fmts_f_t4,na_rep="-",thousands=","))
                        except Exception as e_fmt_f_t4:st.error(f"Err format affichage: {e_fmt_f_t4}");st.dataframe(df_f_disp_t4)
                        st.metric(label="💰 Mnt Total Prévisionnel (€) (mois sel.)",value=f"{gt_f_disp_t4:,.2f} €")
                        st.markdown("#### Export Simulation")
                        out_b_f_t4_exp=io.BytesIO();df_e_f_t4_exp=df_f_disp_t4.copy()
                        try:
                            sim_t_fn_t4=sim_t_f_ui_t4.replace(' ','_').lower()
                            with pd.ExcelWriter(out_b_f_t4_exp,engine="openpyxl")as w_f_t4:
                                sheet_name_fcst_t4 = sanitize_sheet_name(f"Forecast_{sim_t_fn_t4}")
                                df_e_f_t4_exp.to_excel(w_f_t4,sheet_name=sheet_name_fcst_t4,index=False)
                                ws_f_t4=w_f_t4.sheets[sheet_name_fcst_t4]
                                fcst_col_fmts_t4={"Tarif d'achat":"#,##0.00€","Conditionnement":"#,##0"}
                                for m_disp_t4_exp in sel_m_f_ui_t4:
                                    if f"Ventes N-1 {m_disp_t4_exp}"in df_e_f_t4_exp.columns:fcst_col_fmts_t4[f"Ventes N-1 {m_disp_t4_exp}"]="#,##0"
                                    if f"Qté Prév. {m_disp_t4_exp}"in df_e_f_t4_exp.columns:fcst_col_fmts_t4[f"Qté Prév. {m_disp_t4_exp}"]="#,##0"
                                    if f"Montant Prév. {m_disp_t4_exp} (€)"in df_e_f_t4_exp.columns:fcst_col_fmts_t4[f"Montant Prév. {m_disp_t4_exp} (€)"]="#,##0.00€"
                                if"Vts N-1 Tot (Mois Sel.)"in df_e_f_t4_exp.columns:fcst_col_fmts_t4["Vts N-1 Tot (Mois Sel.)"]="#,##0"
                                if"Qté Tot Prév (Mois Sel.)"in df_e_f_t4_exp.columns:fcst_col_fmts_t4["Qté Tot Prév (Mois Sel.)"]="#,##0"
                                if"Mnt Tot Prév (€) (Mois Sel.)"in df_e_f_t4_exp.columns:fcst_col_fmts_t4["Mnt Tot Prév (€) (Mois Sel.)"]="#,##0.00€"
                                format_excel_sheet(ws_f_t4,df_e_f_t4_exp,column_formats=fcst_col_fmts_t4)
                            out_b_f_t4_exp.seek(0)
                            sup_e_nm_f_t4='multi'if len(sel_f_t4)>1 else(sanitize_sheet_name(sel_f_t4[0])if sel_f_t4 else'NA')
                            f_f_exp_t4=f"forecast_{sim_t_fn_t4}_{sup_e_nm_f_t4}_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx"
                            st.download_button("📥 Télécharger Simulation",out_b_f_t4_exp,f_f_exp_t4,"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",key="dl_f_b_t4_dl")
                        except Exception as eef_f_t4:st.error(f"Err export forecast: {eef_f_t4}")
                else:st.info("Paramètres simulation changés. Relancer simulation.")

    # --- Tab 5: Supplier Order Tracking ---
    with tab5:
        # ... (Code identique à la version précédente) ...
        st.header("📄 Suivi des Commandes Fournisseurs")
        if df_suivi_cmds_all is None or df_suivi_cmds_all.empty:
            st.warning("Aucune donnée de suivi (onglet 'Suivi commandes' vide/manquant ou erreur lecture).")
        else:
            sups_in_suivi_list_t5=[]
            if"Fournisseur"in df_suivi_cmds_all.columns:sups_in_suivi_list_t5=sorted(df_suivi_cmds_all["Fournisseur"].astype(str).unique().tolist())
            if not sups_in_suivi_list_t5:st.info("Aucun fournisseur trouvé dans données suivi.")
            else:
                st.markdown("Sélectionnez fournisseurs pour archive de suivi:")
                sel_f_t5_ui = render_supplier_checkboxes("tab5", sups_in_suivi_list_t5, default_select_all=False)
                if not sel_f_t5_ui:st.info("Sélectionner fournisseur(s) pour générer archive suivi.")
                else:
                    st.markdown("---");st.markdown(f"**{len(sel_f_t5_ui)} fournisseur(s) sélectionné(s) pour export.**")
                    if st.button("📦 Générer et Télécharger Archive ZIP de Suivi",key="gen_suivi_zip_btn_t5"):
                        out_cols_s_exp_t5=["Date Pièce BC","N° de pièce","AF_RefFourniss","Désignation Article","Qté Commandées","Date de livraison prévue"]
                        src_cols_need_s_t5=["Date Pièce BC","N° de pièce","AF_RefFourniss","Désignation Article","Qté Commandées","Fournisseur"]
                        miss_src_cols_s_c_t5=[c for c in src_cols_need_s_t5 if c not in df_suivi_cmds_all.columns]
                        if miss_src_cols_s_c_t5:st.error(f"Cols sources manquantes ('Suivi cmds'): {', '.join(miss_src_cols_s_c_t5)}. Export impossible.")
                        else:
                            zip_buf_t5=io.BytesIO();files_added_zip_t5=0
                            try:
                                with zipfile.ZipFile(zip_buf_t5,'w',zipfile.ZIP_DEFLATED)as zipf_t5:
                                    for sup_nm_s_exp_t5 in sel_f_t5_ui:
                                        df_sup_s_exp_d_t5=df_suivi_cmds_all[df_suivi_cmds_all["Fournisseur"]==sup_nm_s_exp_t5].copy()
                                        if df_sup_s_exp_d_t5.empty:logging.info(f"Aucune cmd pour {sup_nm_s_exp_t5}, non ajouté ZIP.");continue
                                        df_exp_fin_s_t5=pd.DataFrame(columns=out_cols_s_exp_t5)
                                        if 'Date Pièce BC' in df_sup_s_exp_d_t5:df_exp_fin_s_t5["Date Pièce BC"]=pd.to_datetime(df_sup_s_exp_d_t5["Date Pièce BC"],errors='coerce')
                                        for col_map_t5 in ["N° de pièce","AF_RefFourniss","Désignation Article","Qté Commandées"]:
                                            if col_map_t5 in df_sup_s_exp_d_t5:df_exp_fin_s_t5[col_map_t5]=df_sup_s_exp_d_t5[col_map_t5]
                                        df_exp_fin_s_t5["Date de livraison prévue"]=""
                                        excel_buf_ind_t5=io.BytesIO()
                                        with pd.ExcelWriter(excel_buf_ind_t5,engine="openpyxl")as writer_ind_t5:
                                            cols_to_write_suivi = [c for c in out_cols_s_exp_t5 if c in df_exp_fin_s_t5.columns]
                                            df_to_w_t5=df_exp_fin_s_t5[cols_to_write_suivi].copy()
                                            sheet_nm_t5=sanitize_sheet_name(f"Suivi_{sup_nm_s_exp_t5}")
                                            df_to_w_t5.to_excel(writer_ind_t5,sheet_name=sheet_nm_t5,index=False)
                                            ws_t5=writer_ind_t5.sheets[sheet_nm_t5]
                                            suivi_col_fmts_t5={"Date Pièce BC":"dd/mm/yyyy","Qté Commandées":"#,##0"}
                                            format_excel_sheet(ws_t5,df_to_w_t5,column_formats=suivi_col_fmts_t5)
                                        excel_b_t5=excel_buf_ind_t5.getvalue()
                                        file_nm_in_zip_t5=f"Suivi_Commande_{sanitize_sheet_name(sup_nm_s_exp_t5)}_{pd.Timestamp.now():%Y%m%d}.xlsx"
                                        zipf_t5.writestr(file_nm_in_zip_t5,excel_b_t5)
                                        files_added_zip_t5+=1
                                if files_added_zip_t5>0:
                                    zip_buf_t5.seek(0)
                                    archive_nm_t5=f"Archive_Suivi_Commandes_{pd.Timestamp.now():%Y%m%d_%H%M}.zip"
                                    st.download_button(label=f"📥 Télécharger Archive ZIP ({files_added_zip_t5} fichier(s))",data=zip_buf_t5,file_name=archive_nm_t5,mime="application/zip",key="dl_suivi_zip_btn_t5_dl")
                                    st.success(f"{files_added_zip_t5} fichier(s) inclus dans ZIP.")
                                else:st.info("Aucun fichier suivi généré (aucun fournisseur sélectionné avec données).")
                            except Exception as e_zip_t5:logging.exception(f"Err création ZIP suivi: {e_zip_t5}");st.error(f"Err création ZIP: {e_zip_t5}")

    # --- Tab 6: New Articles Search ---
    with tab6:
        # ... (Code identique à la version précédente) ...
        st.header("🔍 Recherche des Nouveaux Articles Créés")

        if "Date Création Article" not in st.session_state.df_full.columns:
            st.warning("La colonne 'Date Création Article' est nécessaire et n'a pas été trouvée dans votre fichier. Cette fonctionnalité est désactivée.")
        else:
            if not pd.api.types.is_datetime64_any_dtype(st.session_state.df_full["Date Création Article"]):
                st.error("La colonne 'Date Création Article' n'est pas dans un format de date valide après chargement. Veuillez vérifier le fichier.")
            else:
                min_date_possible = st.session_state.df_full["Date Création Article"].min()
                max_date_possible = st.session_state.df_full["Date Création Article"].max()

                if pd.isna(min_date_possible) or pd.isna(max_date_possible):
                    st.warning("Impossible de déterminer la plage de dates de création. Vérifiez les données.")
                else:
                    default_start_date = max(min_date_possible, pd.Timestamp.now() - pd.DateOffset(months=1))
                    start_date = st.date_input(
                        "Afficher les articles créés à partir du :",
                        value=default_start_date.date(), 
                        min_value=min_date_possible.date(),
                        max_value=max_date_possible.date(),
                        key="new_article_start_date"
                    )
                    if start_date:
                        start_datetime = pd.to_datetime(start_date)
                        source_df_for_new_articles = st.session_state.df_initial_filtered if not st.session_state.df_initial_filtered.empty else st.session_state.df_full
                        if "Date Création Article" in source_df_for_new_articles.columns:
                            df_to_filter = source_df_for_new_articles.copy()
                            if not pd.api.types.is_datetime64_any_dtype(df_to_filter["Date Création Article"]):
                                df_to_filter.loc[:, "Date Création Article"] = pd.to_datetime(df_to_filter["Date Création Article"], errors='coerce')

                            valid_dates_mask = df_to_filter["Date Création Article"].notna()
                            new_articles_df = df_to_filter[valid_dates_mask & (df_to_filter["Date Création Article"] >= start_datetime)].copy()
                            st.markdown(f"--- \n ### {len(new_articles_df)} Nouveaux Articles Trouvés")
                            if not new_articles_df.empty:
                                cols_to_display = ["Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article", "Date Création Article", "Stock", "Tarif d'achat"]
                                existing_cols_to_display = [col for col in cols_to_display if col in new_articles_df.columns]
                                df_display_new = new_articles_df[existing_cols_to_display].copy()
                                if "Date Création Article" in df_display_new.columns:
                                    df_display_new.loc[:, "Date Création Article"] = df_display_new["Date Création Article"].dt.strftime('%d/%m/%Y')
                                st.dataframe(df_display_new)
                                st.markdown("#### Exporter la Liste des Nouveaux Articles")
                                cols_to_export = ["AF_RefFourniss", "Référence Article", "Désignation Article", "Date Création Article"]
                                existing_cols_to_export = [col for col in cols_to_export if col in new_articles_df.columns]
                                if not existing_cols_to_export: st.warning("Colonnes nécessaires à l'export non trouvées.")
                                else:
                                    df_export_new_articles = new_articles_df[existing_cols_to_export].copy()
                                    if "Date Création Article" in df_export_new_articles.columns:
                                        df_export_new_articles.loc[:, "Date Création Article"] = df_export_new_articles["Date Création Article"].dt.strftime('%d/%m/%Y')
                                    output_buffer_new = io.BytesIO()
                                    try:
                                        with pd.ExcelWriter(output_buffer_new, engine="openpyxl") as writer_new:
                                            sheet_name_new = sanitize_sheet_name(f"Nouveaux_Articles_depuis_{start_date.strftime('%Y%m%d')}")
                                            df_export_new_articles.to_excel(writer_new, sheet_name=sheet_name_new, index=False)
                                            ws_new = writer_new.sheets[sheet_name_new]
                                            new_article_fmts = {} 
                                            format_excel_sheet(ws_new, df_export_new_articles, column_formats=new_article_fmts)
                                        output_buffer_new.seek(0)
                                        file_name_new = f"nouveaux_articles_depuis_{start_date.strftime('%Y%m%d')}_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx"
                                        st.download_button(
                                            label=f"📥 Télécharger Nouveaux Articles ({len(df_export_new_articles)} lignes)",
                                            data=output_buffer_new, file_name=file_name_new,
                                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                            key="dl_new_articles_btn_tab6"
                                        )
                                    except Exception as e_export_new: logging.exception(f"Erreur export nouveaux articles: {e_export_new}"); st.error("Erreur création fichier Excel nouveaux articles.")
                            else: st.info("Aucun nouvel article trouvé pour la période sélectionnée.")
                        else: st.error("Colonne 'Date Création Article' non utilisable. Vérifiez le fichier.")

elif not uploaded_file:
    st.info("👋 Bienvenue ! Chargez votre fichier Excel principal pour démarrer.")
    if st.button("🔄 Réinitialiser l'Application"):
        for k_reset in list(st.session_state.keys()): del st.session_state[k_reset]
        for key_reinit, val_reinit in get_default_session_state().items(): st.session_state[key_reinit] = val_reinit
        st.rerun()
elif 'df_initial_filtered' in st.session_state and not isinstance(st.session_state.df_initial_filtered, pd.DataFrame):
    st.error("Erreur interne : Données filtrées invalides. Veuillez recharger le fichier.")
    st.session_state.df_full = None
    if st.button("Réessayer de charger"): st.rerun()

# --- END OF FINAL COMPLETE app.py ---
