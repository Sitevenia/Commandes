import streamlit as st
import pandas as pd
import numpy as np
import io
import logging
import re
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment # For Excel formatting
import calendar
import zipfile # For ZIP export

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
        
        if df is None:
            logging.error(f"Pandas read_excel returned None for sheet '{sheet_name}'.")
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

def format_excel_sheet(worksheet, df, column_formats={}, freeze_header=True, default_float_format="#,##0.00", default_int_format="#,##0", default_date_format="dd/mm/yyyy"):
    """Applies formatting to an openpyxl worksheet based on a DataFrame."""
    if df is None or df.empty:
        logging.warning("Attempted to format sheet with empty or None DataFrame.")
        return

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    data_alignment = Alignment(vertical="center")

    for cell in worksheet[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment

    for idx, col_name in enumerate(df.columns):
        col_letter = get_column_letter(idx + 1)
        number_format_to_apply = None

        try:
            header_len = len(str(col_name))
            non_na_series = df[col_name].dropna()
            sampled_data = non_na_series.sample(min(len(non_na_series), 20)) if not non_na_series.empty else pd.Series([])
            data_len = sampled_data.astype(str).map(len).max() if not sampled_data.empty else 0
            max_len = max(header_len, data_len if pd.notna(data_len) else 0) + 3
            max_len = min(max(max_len, 10), 50)
            worksheet.column_dimensions[col_letter].width = max_len
        except Exception as e:
            logging.warning(f"Could not set width for column {col_name}: {e}")
            worksheet.column_dimensions[col_letter].width = 15

        specific_format = column_formats.get(col_name)
        try: col_dtype = df[col_name].dtype
        except KeyError: logging.warning(f"Column '{col_name}' not in DataFrame for formatting."); continue

        if specific_format: number_format_to_apply = specific_format
        elif pd.api.types.is_integer_dtype(col_dtype): number_format_to_apply = default_int_format
        elif pd.api.types.is_float_dtype(col_dtype): number_format_to_apply = default_float_format
        elif pd.api.types.is_datetime64_any_dtype(col_dtype) or \
             (not df[col_name].empty and isinstance(df[col_name].dropna().iloc[0] if not df[col_name].dropna().empty else None, pd.Timestamp)):
             number_format_to_apply = default_date_format
        
        # Apply format to data rows (worksheet.max_row is based on what df.to_excel wrote)
        for row_idx in range(2, worksheet.max_row + 1): # worksheet.max_row from the df written by to_excel
            cell = worksheet[f"{col_letter}{row_idx}"]
            cell.alignment = data_alignment
            if number_format_to_apply and not str(cell.value).startswith('='):
                try: cell.number_format = number_format_to_apply
                except Exception as e_fmt_cell: logging.warning(f"Could not apply format to cell {col_letter}{row_idx}: {e_fmt_cell}")

    if freeze_header: worksheet.freeze_panes = worksheet['A2']


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
        df_rotation["Semaines Stock (WoS)"] = np.divide(df_rotation["Stock"], denom_wos, out=np.full_like(df_rotation["Stock"], np.inf, dtype=np.float64), where=denom_wos != 0)
        df_rotation.loc[df_rotation["Stock"] <= 0, "Semaines Stock (WoS)"] = 0.0

        denom_rot_unit = df_rotation["Stock"]
        df_rotation["Rotation Unités (Proxy)"] = np.divide(df_rotation["Unités Vendues (Période)"], denom_rot_unit, out=np.full_like(denom_rot_unit, np.inf, dtype=np.float64), where=denom_rot_unit != 0)
        df_rotation.loc[(df_rotation["Unités Vendues (Période)"] <= 0) & (denom_rot_unit <= 0), "Rotation Unités (Proxy)"] = 0.0
        df_rotation.loc[(df_rotation["Unités Vendues (Période)"] <= 0) & (denom_rot_unit > 0), "Rotation Unités (Proxy)"] = 0.0

        df_rotation["COGS (Période)"] = df_rotation["Unités Vendues (Période)"] * df_rotation["Tarif d'achat"]
        df_rotation["Valeur Stock Actuel (€)"] = df_rotation["Stock"] * df_rotation["Tarif d'achat"]
        
        denom_rot_val = df_rotation["Valeur Stock Actuel (€)"]
        df_rotation["Rotation Valeur (Proxy)"] = np.divide(df_rotation["COGS (Période)"], denom_rot_val, out=np.full_like(denom_rot_val, np.inf, dtype=np.float64), where=denom_rot_val != 0)
        df_rotation.loc[(df_rotation["COGS (Période)"] <= 0) & (denom_rot_val <= 0), "Rotation Valeur (Proxy)"] = 0.0
        df_rotation.loc[(df_rotation["COGS (Période)"] <= 0) & (denom_rot_val > 0), "Rotation Valeur (Proxy)"] = 0.0
        return df_rotation
    except KeyError as e: st.error(f"Erreur de clé (rotation): '{e}'."); logging.exception(f"KeyError in calc_rotation: {e}"); return None
    except Exception as e: st.error(f"Erreur inattendue (rotation): {type(e).__name__} - {e}"); logging.exception("Error in calc_rotation:"); return None

def approx_weeks_to_months(week_columns_52):
    month_map = {}
    if not week_columns_52 or len(week_columns_52) != 52: logging.warning(f"approx_weeks_to_months expects 52 cols, got {len(week_columns_52) if week_columns_52 else 0}."); return month_map
    weeks_per_month_approx = 52 / 12.0
    for i in range(1, 13):
        month_name = calendar.month_name[i]
        start_idx = int(round((i-1) * weeks_per_month_approx))
        end_idx = int(round(i * weeks_per_month_approx))
        month_map[month_name] = week_columns_52[start_idx : min(end_idx, 52)]
    logging.info(f"Approx month map. Jan: {month_map.get('January', [])}"); return month_map

def calculer_forecast_simulation_v3(df, all_semaine_columns, selected_months, sim_type, progression_pct=0, objectif_montant=0):
    try:
        if not isinstance(df,pd.DataFrame)or df.empty: st.warning("Aucune donnée pour simu forecast."); return None,0.0
        if not all_semaine_columns or len(all_semaine_columns)<52: st.error("Données histo. < 52 sem. pour N-1."); return None,0.0
        if not selected_months: st.warning("Sélectionner au moins un mois pour simu."); return None,0.0
        required_cols=["Référence Article","Désignation Article","Conditionnement","Tarif d'achat","Fournisseur"]
        if not all(c in df.columns for c in required_cols): st.error(f"Cols manquantes (simu): {', '.join([c for c in required_cols if c not in df.columns])}"); return None,0.0
        
        years_in_cols=set(); parsed_week_cols=[]
        for col_name in all_semaine_columns:
            if isinstance(col_name,str):
                match=re.match(r"(\d{4})S?(\d{1,2})",col_name,re.IGNORECASE)
                if match:
                    year,week=int(match.group(1)),int(match.group(2))
                    if 1<=week<=53: years_in_cols.add(year); parsed_week_cols.append({'year':year,'week':week,'col':col_name,'sort_key':year*100+week})
        if not years_in_cols: st.error("Impossible de déterminer années. Format: 'YYYYWW' ou 'YYYYSwW'."); return None,0.0
        parsed_week_cols.sort(key=lambda x:x['sort_key'])
        year_n=max(years_in_cols)if years_in_cols else 0; year_n_minus_1=year_n-1
        st.caption(f"Simu N-1 (N: {year_n}, N-1: {year_n_minus_1})")
        n1_week_cols_data=[item for item in parsed_week_cols if item['year']==year_n_minus_1]
        if len(n1_week_cols_data)<52: st.error(f"Données N-1 ({year_n_minus_1}) < 52 sem. ({len(n1_week_cols_data)})."); return None,0.0
        n1_week_cols_for_mapping=[item['col']for item in n1_week_cols_data[:52]]

        df_sim=df[required_cols].copy()
        df_sim["Tarif d'achat"]=pd.to_numeric(df_sim["Tarif d'achat"],errors='coerce').fillna(0)
        df_sim["Conditionnement"]=pd.to_numeric(df_sim["Conditionnement"],errors='coerce').fillna(1).apply(lambda x:1 if x<=0 else int(x))
        if not all(c in df.columns for c in n1_week_cols_for_mapping): st.error(f"Err interne: Cols N-1 mappées non trouvées."); return None,0.0
        df_n1_sales_data=df[n1_week_cols_for_mapping].copy()
        for col in n1_week_cols_for_mapping:df_n1_sales_data[col]=pd.to_numeric(df_n1_sales_data[col],errors='coerce').fillna(0)
        
        month_col_map_n1=approx_weeks_to_months(n1_week_cols_for_mapping)
        total_n1_sales_selected_months_series=pd.Series(0.0,index=df_sim.index)
        monthly_sales_n1_for_selected_months={}
        for month_name in selected_months:
            sales_this_month_n1=pd.Series(0.0,index=df_sim.index)
            if month_name in month_col_map_n1 and month_col_map_n1[month_name]:
                actual_cols=[c for c in month_col_map_n1[month_name]if c in df_n1_sales_data.columns] # Shortened
                if actual_cols:sales_this_month_n1=df_n1_sales_data[actual_cols].sum(axis=1)
            monthly_sales_n1_for_selected_months[month_name]=sales_this_month_n1
            total_n1_sales_selected_months_series+=sales_this_month_n1
            df_sim[f"Ventes N-1 {month_name}"]=sales_this_month_n1
        df_sim["Vts N-1 Tot (Mois Sel.)"]=total_n1_sales_selected_months_series

        period_seasonality_factors={}
        safe_total_n1=total_n1_sales_selected_months_series.copy() # Shortened
        for month_name in selected_months:
            month_sales_n1=monthly_sales_n1_for_selected_months.get(month_name,pd.Series(0.0,index=df_sim.index))
            factor=np.divide(month_sales_n1,safe_total_n1,out=np.zeros_like(month_sales_n1,dtype=float),where=safe_total_n1!=0)
            period_seasonality_factors[month_name]=pd.Series(factor,index=df_sim.index).fillna(0)

        base_monthly_forecast_qty_map={}
        if sim_type=='Simple Progression':
            prog_factor=1+(progression_pct/100.0)
            total_fcst_qty_period=total_n1_sales_selected_months_series*prog_factor # Shortened
            for m_name in selected_months: # Shortened
                seas_factor=period_seasonality_factors.get(m_name,pd.Series(0.0,index=df_sim.index)) # Shortened
                base_monthly_forecast_qty_map[m_name]=total_fcst_qty_period*seas_factor
        elif sim_type=='Objectif Montant':
            if objectif_montant<=0:st.error("Objectif Montant > 0 requis.");return None,0.0
            total_n1_units_all=total_n1_sales_selected_months_series.sum() # Shortened
            if total_n1_units_all<=0:
                st.warning("Ventes N-1 nulles. Répartition égale du montant objectif.")
                num_sel_m=len(selected_months);if num_sel_m==0:return None,0.0 # Shortened
                target_amt_p_m=objectif_montant/num_sel_m # Shortened
                num_items_price=(df_sim["Tarif d'achat"]>0).sum() # Shortened
                for m_name in selected_months:
                    if num_items_price==0:base_monthly_forecast_qty_map[m_name]=pd.Series(0.0,index=df_sim.index)
                    else:
                        target_amt_p_item_m=target_amt_p_m/num_items_price # Shortened
                        base_monthly_forecast_qty_map[m_name]=np.divide(target_amt_p_item_m,df_sim["Tarif d'achat"],out=np.zeros_like(df_sim["Tarif d'achat"],dtype=float),where=df_sim["Tarif d'achat"]!=0)
            else:
                for m_name in selected_months:
                    seas_factor=period_seasonality_factors.get(m_name,pd.Series(0.0,index=df_sim.index))
                    target_amt_m_item=objectif_montant*seas_factor # Shortened
                    base_monthly_forecast_qty_map[m_name]=np.divide(target_amt_m_item,df_sim["Tarif d'achat"],out=np.zeros_like(df_sim["Tarif d'achat"],dtype=float),where=df_sim["Tarif d'achat"]!=0)
        else:st.error(f"Type simu non reconnu: '{sim_type}'.");return None,0.0

        tot_adj_qty_all_m=pd.Series(0.0,index=df_sim.index) # Shortened
        tot_fin_amt_all_m=pd.Series(0.0,index=df_sim.index) # Shortened
        for m_name in selected_months:
            fcst_qty_col,fcst_amt_col=f"Qté Prév. {m_name}",f"Montant Prév. {m_name} (€)" # Shortened
            base_q_s=base_monthly_forecast_qty_map.get(m_name,pd.Series(0.0,index=df_sim.index)) # Shortened
            base_q_s=pd.to_numeric(base_q_s,errors='coerce').fillna(0)
            cond_s=df_sim["Conditionnement"] # Shortened
            adj_qty_s=(np.ceil(np.divide(base_q_s,cond_s,out=np.zeros_like(base_q_s,dtype=float),where=cond_s!=0))*cond_s).fillna(0).astype(int) # Shortened
            df_sim[fcst_qty_col]=adj_qty_s;df_sim[fcst_amt_col]=adj_qty_s*df_sim["Tarif d'achat"]
            tot_adj_qty_all_m+=adj_qty_s;tot_fin_amt_all_m+=df_sim[fcst_amt_col]
        
        df_sim["Qté Totale Prév. (Mois Sel.)"]=tot_adj_qty_all_m
        df_sim["Montant Total Prév. (€) (Mois Sel.)"]=tot_fin_amt_all_m
        id_cols_d=["Fournisseur","Référence Article","Désignation Article","Conditionnement","Tarif d'achat"] # Shortened
        n1_sales_cols_d=sorted([f"Ventes N-1 {m}"for m in selected_months if f"Ventes N-1 {m}"in df_sim.columns]) # Shortened
        qty_fcst_cols_d=sorted([f"Qté Prév. {m}"for m in selected_months if f"Qté Prév. {m}"in df_sim.columns]) # Shortened
        amt_fcst_cols_d=sorted([f"Montant Prév. {m} (€)"for m in selected_months if f"Montant Prév. {m} (€)"in df_sim.columns]) # Shortened
        df_sim.rename(columns={"Qté Totale Prév. (Mois Sel.)":"Qté Tot Prév (Mois Sel.)","Montant Total Prév. (€) (Mois Sel.)":"Mnt Tot Prév (€) (Mois Sel.)"},inplace=True)
        total_cols_d=["Vts N-1 Tot (Mois Sel.)","Qté Tot Prév (Mois Sel.)","Mnt Tot Prév (€) (Mois Sel.)"] # Shortened
        final_ord_cols=id_cols_d+total_cols_d+n1_sales_cols_d+qty_fcst_cols_d+amt_fcst_cols_d # Shortened
        final_ord_cols_exist=[c for c in final_ord_cols if c in df_sim.columns] # Shortened
        grand_total_fcst_amt=tot_fin_amt_all_m.sum() # Shortened
        return df_sim[final_ord_cols_exist],grand_total_fcst_amt
    except KeyError as e:st.error(f"Err clé (simu fcst): '{e}'.");logging.exception(f"KeyError in calc_fcst_sim_v3: {e}");return None,0.0 # Shortened
    except Exception as e:st.error(f"Err inattendue (simu fcst): {type(e).__name__} - {e}");logging.exception("Error in calc_fcst_sim_v3:");return None,0.0 # Shortened

def sanitize_sheet_name(name):
    if not isinstance(name,str):name=str(name)
    s=re.sub(r'[\[\]:*?/\\<>|"]','_',name); # Shortened
    if s.startswith("'"):s="_"+s[1:]
    if s.endswith("'"):s=s[:-1]+"_"
    return s[:31]

def render_supplier_checkboxes(tab_key_prefix,all_suppliers,default_select_all=False):
    sel_all_k=f"{tab_key_prefix}_select_all" # Shortened
    sup_cb_ks={s:f"{tab_key_prefix}_cb_{sanitize_supplier_key(s)}"for s in all_suppliers} # Shortened
    if sel_all_k not in st.session_state:
        st.session_state[sel_all_k]=default_select_all
        for cb_k in sup_cb_ks.values():
            if cb_k not in st.session_state:st.session_state[cb_k]=default_select_all
    else:
        for cb_k in sup_cb_ks.values():
            if cb_k not in st.session_state:st.session_state[cb_k]=st.session_state[sel_all_k]
    def toggle_all_s_tab(): # Shortened
        curr_sel_all_v=st.session_state[sel_all_k] # Shortened
        for cb_k in sup_cb_ks.values():st.session_state[cb_k]=curr_sel_all_v
    def check_ind_s_tab(): # Shortened
        all_ind_chk=all(st.session_state.get(cb_k,False)for cb_k in sup_cb_ks.values()) # Shortened
        if st.session_state.get(sel_all_k)!=all_ind_chk:st.session_state[sel_all_k]=all_ind_chk
    exp_lbl="👤 Sélectionner Fournisseurs" # Shortened
    if tab_key_prefix=="tab5":exp_lbl="👤 Sélectionner Fournisseurs pour Export Suivi"
    with st.expander(exp_lbl,expanded=True):
        st.checkbox("Sélectionner / Désélectionner Tout",key=sel_all_k,on_change=toggle_all_s_tab,disabled=not bool(all_suppliers))
        st.markdown("---")
        sel_sups_ui=[] # Shortened
        num_disp_cols=4;chk_cols=st.columns(num_disp_cols);curr_col_idx=0 # Shortened
        for sup_n,cb_k in sup_cb_ks.items(): # Shortened
            chk_cols[curr_col_idx].checkbox(sup_n,key=cb_k,on_change=check_ind_s_tab)
            if st.session_state.get(cb_k):sel_sups_ui.append(sup_n)
            curr_col_idx=(curr_col_idx+1)%num_disp_cols
    return sel_sups_ui

def sanitize_supplier_key(supplier_name):
    if not isinstance(supplier_name,str):supplier_name=str(supplier_name)
    s=re.sub(r'\W+','_',supplier_name);s=re.sub(r'^_+|_+$','',s);s=re.sub(r'_+','_',s)
    return s if s else"invalid_supplier_key"

st.set_page_config(page_title="Forecast & Rotation App",layout="wide")
st.title("📦 Application Prévision Commande, Analyse Rotation & Suivi")
uploaded_file=st.file_uploader("📁 Charger le fichier Excel principal",type=["xlsx","xls"],key="main_file_uploader")

def get_default_session_state():
    return {'df_full':None,'min_order_dict':{},'df_initial_filtered':pd.DataFrame(),'all_available_semaine_columns':[],'unique_suppliers_list':[],'commande_result_df':None,'commande_calculated_total_amount':0.0,'commande_suppliers_calculated_for':[],'rotation_result_df':None,'rotation_analysis_period_label':"",'rotation_suppliers_calculated_for':[],'rotation_threshold_value':1.0,'show_all_rotation_data':True,'forecast_result_df':None,'forecast_grand_total_amount':0.0,'forecast_simulation_params_calculated_for':{},'forecast_selected_months_ui':list(calendar.month_name)[1:],'forecast_sim_type_radio_index':0,'forecast_progression_percentage_ui':5.0,'forecast_target_amount_ui':10000.0,'df_suivi_commandes':None}
for k,v_def in get_default_session_state().items(): # Shortened
    if k not in st.session_state:st.session_state[k]=v_def

if uploaded_file and st.session_state.df_full is None:
    logging.info(f"New file: {uploaded_file.name}. Processing...")
    keys_to_reset=list(get_default_session_state().keys())
    dyn_key_prefs=['tab1_','tab2_','tab3_','tab4_','tab5_'] # Shortened
    for k in keys_to_reset:
        if k in st.session_state:del st.session_state[k]
    for pref in dyn_key_prefs: # Shortened
        for k_rem in[k for k in st.session_state if k.startswith(pref)]:del st.session_state[k_rem] # Shortened
    for k,v_def in get_default_session_state().items():st.session_state[k]=v_def
    logging.info("Session state reset for new file.")
    try:
        excel_buf=io.BytesIO(uploaded_file.getvalue()) # Shortened
        st.info("Lecture 'Tableau final'...")
        df_full_t=safe_read_excel(excel_buf,sheet_name="Tableau final",header=7) # Shortened
        if df_full_t is None:st.error("❌ Échec lecture 'TF'.");st.stop() # Shortened
        req_tf_cols=["Stock","Fournisseur","AF_RefFourniss","Tarif d'achat","Conditionnement","Référence Article","Désignation Article"] # Shortened
        if not all(c in df_full_t.columns for c in req_tf_cols):st.error(f"❌ Cols manquantes ('TF'): {', '.join([c for c in req_tf_cols if c not in df_full_t.columns])}");st.stop()
        df_full_t["Stock"]=pd.to_numeric(df_full_t["Stock"],errors='coerce').fillna(0)
        df_full_t["Tarif d'achat"]=pd.to_numeric(df_full_t["Tarif d'achat"],errors='coerce').fillna(0)
        df_full_t["Conditionnement"]=pd.to_numeric(df_full_t["Conditionnement"],errors='coerce').fillna(1).apply(lambda x:int(x)if x>0 else 1)
        for str_c in["Fournisseur","AF_RefFourniss","Référence Article","Désignation Article"]: # Shortened
            if str_c in df_full_t.columns:df_full_t[str_c]=df_full_t[str_c].astype(str).str.strip()
        st.session_state.df_full=df_full_t;st.success("✅ 'TF' lu.")
        st.info("Lecture 'Min commande'...") # Shortened
        excel_buf.seek(0)
        df_min_c_t=safe_read_excel(excel_buf,sheet_name="Minimum de commande") # Shortened
        min_o_dict_t={} # Shortened
        if df_min_c_t is not None:
            s_c,m_c="Fournisseur","Minimum de Commande"
            if s_c in df_min_c_t.columns and m_c in df_min_c_t.columns:
                try:
                    df_min_c_t[s_c]=df_min_c_t[s_c].astype(str).str.strip()
                    df_min_c_t[m_c]=pd.to_numeric(df_min_c_t[m_c],errors='coerce')
                    min_o_dict_t=df_min_c_t.dropna(subset=[s_c,m_c]).set_index(s_c)[m_c].to_dict()
                    st.success(f"✅ 'Min cmd' lu ({len(min_o_dict_t)}).") # Shortened
                except Exception as e_min:st.error(f"❌ Err trait. 'Min cmd': {e_min}") # Shortened
            else:st.warning(f"⚠️ Cols '{s_c}'/'{m_c}' manquantes ('Min cmd').") # Shortened
        st.session_state.min_order_dict=min_o_dict_t
        st.info("Lecture 'Suivi commandes'...")
        excel_buf.seek(0)
        df_suivi_t=safe_read_excel(excel_buf,sheet_name="Suivi commandes",header=4) # Shortened
        if df_suivi_t is not None:
            req_s_cols=["Date Pièce BC","N° de pièce","AF_RefFourniss","Désignation Article","Qté Commandées","Intitulé Fournisseur"] # Shortened
            miss_s_cols_c=[c for c in req_s_cols if c not in df_suivi_t.columns] # Shortened
            if not miss_s_cols_c:
                df_suivi_t.rename(columns={"Intitulé Fournisseur":"Fournisseur"},inplace=True)
                for col_strp in["Fournisseur","AF_RefFourniss","Désignation Article","N° de pièce"]: # Shortened
                    if col_strp in df_suivi_t.columns:df_suivi_t[col_strp]=df_suivi_t[col_strp].astype(str).str.strip()
                if "Qté Commandées"in df_suivi_t.columns:df_suivi_t["Qté Commandées"]=pd.to_numeric(df_suivi_t["Qté Commandées"],errors='coerce').fillna(0)
                if "Date Pièce BC"in df_suivi_t.columns:
                    try:df_suivi_t["Date Pièce BC"]=pd.to_datetime(df_suivi_t["Date Pièce BC"],errors='coerce')
                    except Exception as e_dt:st.warning(f"⚠️ Problème parsing 'Date Pièce BC': {e_dt}.")
                df_suivi_t.dropna(how='all',inplace=True)
                st.session_state.df_suivi_commandes=df_suivi_t
                st.success(f"✅ 'Suivi cmds' lu ({len(df_suivi_t)}).") # Shortened
            else:
                st.warning(f"⚠️ Cols manquantes ('Suivi cmds', L5): {', '.join(miss_s_cols_c)}. Suivi limité.")
                st.session_state.df_suivi_commandes=pd.DataFrame()
        else:
            st.info("Onglet 'Suivi cmds' non trouvé/vide. Suivi non dispo.") # Shortened
            st.session_state.df_suivi_commandes=pd.DataFrame()
        df_ld_ff=st.session_state.df_full # Shortened
        df_init_filt_t=df_ld_ff[(df_ld_ff["Fournisseur"].notna())&(df_ld_ff["Fournisseur"]!="")&(df_ld_ff["Fournisseur"]!="#FILTER")&(df_ld_ff["AF_RefFourniss"].notna())&(df_ld_ff["AF_RefFourniss"]!="")].copy() # Shortened
        st.session_state.df_initial_filtered=df_init_filt_t
        f_w_c_idx=12;pot_s_cols=[] # Shortened
        if len(df_ld_ff.columns)>f_w_c_idx:
            cand_c_s=df_ld_ff.columns[f_w_c_idx:].tolist() # Shortened
            known_non_w_c=["Tarif d'achat","Conditionnement","Stock","Total","Stock à terme","Ventes N-1","Ventes 12 semaines identiques N-1","Ventes 12 dernières semaines","Quantité à commander","Fournisseur","AF_RefFourniss","Référence Article","Désignation Article"] # Shortened
            excl_s=set(known_non_w_c) # Shortened
            for col_c in cand_c_s:
                if col_c not in excl_s and pd.api.types.is_numeric_dtype(df_ld_ff.get(col_c,pd.Series(dtype=object)).dtype):pot_s_cols.append(col_c)
        st.session_state.all_available_semaine_columns=pot_s_cols
        if not pot_s_cols:st.warning("⚠️ Aucune col vente numérique identifiée.") # Shortened
        if not df_init_filt_t.empty:st.session_state.unique_suppliers_list=sorted(df_init_filt_t["Fournisseur"].astype(str).unique().tolist())
        st.rerun()
    except Exception as e_load_main:
        st.error(f"❌ Err majeure chargement/traitement: {e_load_main}") # Shortened
        logging.exception("Major file loading/processing error:")
        st.session_state.df_full=None;st.stop()

if 'df_initial_filtered'in st.session_state and isinstance(st.session_state.df_initial_filtered,pd.DataFrame):
    df_base_tabs=st.session_state.df_initial_filtered # Shortened
    all_sups_data=st.session_state.unique_suppliers_list # Shortened
    min_o_amts=st.session_state.min_order_dict # Shortened
    id_sem_cols=st.session_state.all_available_semaine_columns # Shortened
    df_suivi_cmds_all=st.session_state.get('df_suivi_commandes',pd.DataFrame()) # Shortened

    tab_titles=["Prévision Commande","Analyse Rotation Stock","Vérification Stock","Simulation Forecast","Suivi Commandes Fourn."]
    tab1,tab2,tab3,tab4,tab5=st.tabs(tab_titles)

    with tab1: # Prévision Commande
        st.header("Prévision des Quantités à Commander")
        sel_f_t1=render_supplier_checkboxes("tab1",all_sups_data,default_select_all=True) # Shortened
        df_disp_t1=pd.DataFrame() # Shortened
        if sel_f_t1:
            if not df_base_tabs.empty:
                df_disp_t1=df_base_tabs[df_base_tabs["Fournisseur"].isin(sel_f_t1)].copy()
                st.caption(f"{len(df_disp_t1)} art. / {len(sel_f_t1)} fourn.")
        else:st.info("Sélectionner fournisseur(s).")
        st.markdown("---")
        if df_disp_t1.empty and sel_f_t1:st.warning("Aucun article pour fournisseur(s) sélectionné(s).")
        elif not id_sem_cols and not df_disp_t1.empty:st.warning("Colonnes ventes non identifiées.")
        elif not df_disp_t1.empty:
            st.markdown("#### Paramètres Calcul Commande")
            c1_c,c2_c=st.columns(2); # Shortened
            with c1_c:d_s_c=st.number_input("⏳ Couverture (sem.)",1,260,4,1,key="d_s_c_t1")
            with c2_c:m_m_c=st.number_input("💶 Montant min (€)",0.0,value=0.0,step=50.0,format="%.2f",key="m_m_c_t1")
            if st.button("🚀 Calculer Qtés Cmd",key="calc_q_c_b_t1"):
                with st.spinner("Calcul qtés..."):res_c=calculer_quantite_a_commander(df_disp_t1,id_sem_cols,m_m_c,d_s_c)
                if res_c:
                    st.success("✅ Calcul qtés OK.");q_c,vN1,v12N1,v12l,m_c=res_c
                    df_r_c=df_disp_t1.copy();df_r_c["Qte Cmdée"]=q_c
                    df_r_c["Vts N-1 Total (calc)"]=vN1;df_r_c["Vts 12 N-1 Sim (calc)"]=v12N1;df_r_c["Vts 12 Dern. (calc)"]=v12l
                    df_r_c["Tarif Ach."]=pd.to_numeric(df_r_c["Tarif d'achat"],errors='coerce').fillna(0)
                    df_r_c["Total Cmd (€)"]=df_r_c["Tarif Ach."]*df_r_c["Qte Cmdée"]
                    df_r_c["Stock Terme"]=df_r_c["Stock"]+df_r_c["Qte Cmdée"]
                    st.session_state.commande_result_df=df_r_c;st.session_state.commande_calculated_total_amount=m_c
                    st.session_state.commande_suppliers_calculated_for=sel_f_t1;st.rerun()
                else:st.error("❌ Calcul qtés échoué.")
            if st.session_state.commande_result_df is not None and st.session_state.commande_suppliers_calculated_for==sel_f_t1:
                st.markdown("---");st.markdown("#### Résultats Prévision Commande")
                df_c_d=st.session_state.commande_result_df;m_c_d=st.session_state.commande_calculated_total_amount;s_c_d=st.session_state.commande_suppliers_calculated_for
                st.metric(label="💰 Montant Total Cmd",value=f"{m_c_d:,.2f} €")
                if len(s_c_d)==1:
                    s_s=s_c_d[0]
                    if s_s in min_o_amts:
                        r_m_s=min_o_amts[s_s];a_t_s=df_c_d[df_c_d["Fournisseur"]==s_s]["Total Cmd (€)"].sum()
                        if r_m_s>0 and a_t_s<r_m_s:st.warning(f"⚠️ Min non atteint ({s_s}): {a_t_s:,.2f}€ / Requis: {r_m_s:,.2f}€ (Manque: {r_m_s-a_t_s:,.2f}€)")
                cols_s_c=["Fournisseur","AF_RefFourniss","Référence Article","Désignation Article","Stock","Vts N-1 Total (calc)","Vts 12 N-1 Sim (calc)","Vts 12 Dern. (calc)","Conditionnement","Qte Cmdée","Stock Terme","Tarif Ach.","Total Cmd (€)"]
                disp_c_c=[c for c in cols_s_c if c in df_c_d.columns]
                if not disp_c_c:st.error("Aucune col à afficher (cmd).")
                else:
                    fmts_c={"Tarif Ach.":"{:,.2f}€","Total Cmd (€)":"{:,.2f}€","Vts N-1 Total (calc)":"{:,.0f}","Vts 12 N-1 Sim (calc)":"{:,.0f}","Vts 12 Dern. (calc)":"{:,.0f}","Stock":"{:,.0f}","Conditionnement":"{:,.0f}","Qte Cmdée":"{:,.0f}","Stock Terme":"{:,.0f}"}
                    st.dataframe(df_c_d[disp_c_c].style.format(fmts_c,na_rep="-",thousands=","))
                st.markdown("#### Export Commandes")
                df_e_c=df_c_d[df_c_d["Qte Cmdée"]>0].copy()
                if not df_e_c.empty:
                    out_b_c=io.BytesIO();shts_c=0
                    try:
                        with pd.ExcelWriter(out_b_c,engine="openpyxl") as writer_c:
                            exp_c_s_c=[c for c in disp_c_c if c!='Fournisseur']
                            q,p,t="Qte Cmdée","Tarif Ach.","Total Cmd (€)"
                            f_ok=False
                            if all(c in exp_c_s_c for c in[q,p,t]):
                                try:q_l,p_l,t_l=get_column_letter(exp_c_s_c.index(q)+1),get_column_letter(exp_c_s_c.index(p)+1),get_column_letter(exp_c_s_c.index(t)+1);f_ok=True
                                except ValueError:pass
                            for sup_e in s_c_d:
                                df_s_e=df_e_c[df_e_c["Fournisseur"]==sup_e]
                                if not df_s_e.empty:
                                    df_w_s=df_s_e[exp_c_s_c].copy()
                                    n_r=len(df_w_s);s_nm=sanitize_sheet_name(sup_e)
                                    try:
                                        df_w_s.to_excel(writer_c,sheet_name=s_nm,index=False)
                                        ws=writer_c.sheets[s_nm]
                                        cmd_col_fmts={"Stock":"#,##0","Vts N-1 Total (calc)":"#,##0","Vts 12 N-1 Sim (calc)":"#,##0","Vts 12 Dern. (calc)":"#,##0","Conditionnement":"#,##0","Qte Cmdée":"#,##0","Stock Terme":"#,##0","Tarif Ach.":"#,##0.00€"}
                                        format_excel_sheet(ws,df_w_s,column_formats=cmd_col_fmts)
                                        if f_ok and n_r>0:
                                            for r_idx in range(2,n_r+2):
                                                cell_t=ws[f"{t_l}{r_idx}"];cell_t.value=f"={q_l}{r_idx}*{p_l}{r_idx}";cell_t.number_format='#,##0.00€' # Shortened
                                        lbl_c_s_idx=exp_c_s_c.index("Désignation Article"if"Désignation Article"in exp_c_s_c else(exp_c_s_c[1]if len(exp_c_s_c)>1 else exp_c_s_c[0]))+1
                                        tot_v_s=df_w_s[t].sum();min_r_s=min_o_amts.get(sup_e,0);min_d_s=f"{min_r_s:,.2f}€"if min_r_s>0 else"N/A"
                                        total_row_xl_idx=n_r+2 # Shortened
                                        ws[f"{get_column_letter(lbl_c_s_idx)}{total_row_xl_idx}"]="TOTAL"
                                        ws[f"{get_column_letter(lbl_c_s_idx)}{total_row_xl_idx}"].font=Font(bold=True)
                                        cell_gt=ws[f"{t_l}{total_row_xl_idx}"] # Shortened
                                        if n_r>0:cell_gt.value=f"=SUM({t_l}2:{t_l}{n_r+1})"
                                        else:cell_gt.value=tot_v_s
                                        cell_gt.number_format='#,##0.00€';cell_gt.font=Font(bold=True)
                                        min_req_row_xl_idx=n_r+3 # Shortened
                                        ws[f"{get_column_letter(lbl_c_s_idx)}{min_req_row_xl_idx}"]="Min Requis Fourn."
                                        ws[f"{get_column_letter(lbl_c_s_idx)}{min_req_row_xl_idx}"].font=Font(bold=True)
                                        cell_min_req_v=ws[f"{t_l}{min_req_row_xl_idx}"] # Shortened
                                        cell_min_req_v.value=min_d_s;cell_min_req_v.font=Font(bold=True)
                                        shts_c+=1
                                    except Exception as e_sht:logging.error(f"Err export sheet {s_nm}: {e_sht}")
                        if shts_c>0:
                            out_b_c.seek(0) # Moved writer.save() as it's handled by context manager
                            fn_c=f"commandes_{'multi'if len(s_c_d)>1 else sanitize_sheet_name(s_c_d[0])}_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx"
                            st.download_button(f"📥 Télécharger ({shts_c} feuilles)",out_b_c,fn_c,"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",key="dl_c_b_t1_dl")
                        else:st.info("Aucune qté > 0 à exporter (ou err création feuilles).")
                    except Exception as e_wrt_c:logging.exception(f"Err ExcelWriter cmd: {e_wrt_c}");st.error("Erreur export commandes.")
                else:st.info("Aucun article qté > 0 à exporter.")
            else:st.info("Résultats commande invalidés. Relancer.")

    with tab2: # Analyse Rotation Stock
        st.header("Analyse de la Rotation des Stocks")
        sel_f_t2=render_supplier_checkboxes("tab2",all_sups_data,default_select_all=True) # Shortened
        df_disp_t2=pd.DataFrame() # Shortened
        if sel_f_t2:
            if not df_base_tabs.empty:
                df_disp_t2=df_base_tabs[df_base_tabs["Fournisseur"].isin(sel_f_t2)].copy()
                st.caption(f"{len(df_disp_t2)} art. / {len(sel_f_t2)} fourn.")
        else:st.info("Sélectionner fournisseur(s).")
        st.markdown("---")
        if df_disp_t2.empty and sel_f_t2:st.warning("Aucun article pour fournisseur(s) sélectionné(s).")
        elif not id_sem_cols and not df_disp_t2.empty:st.warning("Colonnes ventes non identifiées.")
        elif not df_disp_t2.empty:
            st.markdown("#### Paramètres Analyse Rotation")
            c1_r,c2_r=st.columns(2); # Shortened
            with c1_r:p_opts_r={"12 dern. sem.":12,"52 dern. sem.":52,"Total dispo.":0};sel_p_lbl_r=st.selectbox("⏳ Période analyse:",p_opts_r.keys(),key="r_p_sel_ui_t2");sel_p_w_r=p_opts_r[sel_p_lbl_r]
            with c2_r:
                st.markdown("##### Options Affichage");show_all_r=st.checkbox("Afficher tout",value=st.session_state.show_all_rotation_data,key="show_all_r_ui_cb_t2");st.session_state.show_all_rotation_data=show_all_r
                r_thr_ui=st.number_input("... ou vts mens. <",0.0,value=st.session_state.rotation_threshold_value,step=0.1,format="%.1f",key="r_thr_ui_numin_t2",disabled=show_all_r)
                if not show_all_r:st.session_state.rotation_threshold_value=r_thr_ui
            if st.button("🔄 Analyser Rotation",key="analyze_r_btn_t2"):
                with st.spinner("Analyse rotation..."):df_r_res=calculer_rotation_stock(df_disp_t2,id_sem_cols,sel_p_w_r)
                if df_r_res is not None:
                    st.success("✅ Analyse rotation OK.");st.session_state.rotation_result_df=df_r_res
                    st.session_state.rotation_analysis_period_label=sel_p_lbl_r;st.session_state.rotation_suppliers_calculated_for=sel_f_t2;st.rerun()
                else:st.error("❌ Analyse rotation échouée.")
            if st.session_state.rotation_result_df is not None and st.session_state.rotation_suppliers_calculated_for==sel_f_t2:
                st.markdown("---");st.markdown(f"#### Résultats Rotation ({st.session_state.rotation_analysis_period_label})")
                df_r_orig=st.session_state.rotation_result_df;thr_d_r=st.session_state.rotation_threshold_value;show_all_f_r=st.session_state.show_all_rotation_data
                m_sales_c_r="Ventes Moy Mensuel (Période)";df_r_disp=pd.DataFrame();df_r_to_fmt=pd.DataFrame() # Shortened
                if df_r_orig.empty:st.info("Aucune donnée rotation à afficher.")
                elif show_all_f_r:df_r_disp=df_r_orig.copy();df_r_to_fmt=df_r_disp.copy();st.caption(f"Affichage {len(df_r_disp)} articles.")
                elif m_sales_c_r in df_r_orig.columns:
                    try:
                        sales_f=pd.to_numeric(df_r_orig[m_sales_c_r],errors='coerce').fillna(0)
                        df_r_disp=df_r_orig[sales_f<thr_d_r].copy();df_r_to_fmt=df_r_disp.copy()
                        st.caption(f"Filtre: Vts < {thr_d_r:.1f}/mois. {len(df_r_disp)} / {len(df_r_orig)} art.")
                        if df_r_disp.empty:st.info(f"Aucun article < {thr_d_r:.1f} vts/mois.")
                    except Exception as ef_r:st.error(f"Err filtre: {ef_r}");df_r_disp=df_r_orig.copy();df_r_to_fmt=df_r_disp.copy()
                else:st.warning(f"Col '{m_sales_c_r}' non trouvée. Affichage tout.");df_r_disp=df_r_orig.copy();df_r_to_fmt=df_r_disp.copy()
                if not df_r_disp.empty:
                    cols_r_s=["Fournisseur","AF_RefFourniss","Référence Article","Désignation Article","Tarif d'achat","Stock","Unités Vendues (Période)","Ventes Moy Hebdo (Période)","Ventes Moy Mensuel (Période)","Semaines Stock (WoS)","Rotation Unités (Proxy)","Valeur Stock Actuel (€)","COGS (Période)","Rotation Valeur (Proxy)"]
                    disp_c_r=[c for c in cols_r_s if c in df_r_disp.columns]
                    df_d_cp_r=df_r_disp[disp_c_r].copy()
                    num_rnd_r={"Tarif d'achat":2,"Ventes Moy Hebdo (Période)":2,"Ventes Moy Mensuel (Période)":2,"Semaines Stock (WoS)":1,"Rotation Unités (Proxy)":2,"Valeur Stock Actuel (€)":2,"COGS (Période)":2,"Rotation Valeur (Proxy)":2}
                    for c,d in num_rnd_r.items():
                        if c in df_d_cp_r.columns:df_d_cp_r[c]=pd.to_numeric(df_d_cp_r[c],errors='coerce').round(d)
                    df_d_cp_r.replace([np.inf,-np.inf],'Infini',inplace=True)
                    fmts_r={"Tarif d'achat":"{:,.2f}€","Stock":"{:,.0f}","Unités Vendues (Période)":"{:,.0f}","Ventes Moy Hebdo (Période)":"{:,.2f}","Ventes Moy Mensuel (Période)":"{:,.2f}","Semaines Stock (WoS)":"{}","Rotation Unités (Proxy)":"{}","Valeur Stock Actuel (€)":"{:,.2f}€","COGS (Période)":"{:,.2f}€","Rotation Valeur (Proxy)":"{}"}
                    st.dataframe(df_d_cp_r.style.format(fmts_r,na_rep="-",thousands=","))
                    st.markdown("#### Export Analyse Affichée")
                    if not df_r_to_fmt.empty:
                        out_b_r=io.BytesIO();df_e_r=df_r_to_fmt[disp_c_r].copy()
                        lbl_e_r=f"Filtree_{thr_d_r:.1f}"if not show_all_f_r else"Complete";sh_nm_r=sanitize_sheet_name(f"Rotation_{lbl_e_r}");f_base_r=f"analyse_rotation_{lbl_e_r}"
                        sup_e_nm_r='multi'if len(sel_f_t2)>1 else(sanitize_sheet_name(sel_f_t2[0])if sel_f_t2 else'NA')
                        try:
                            with pd.ExcelWriter(out_b_r,engine="openpyxl")as wr_r:
                                df_e_r.to_excel(wr_r,sheet_name=sh_nm_r,index=False)
                                ws_r=wr_r.sheets[sh_nm_r]
                                rot_col_fmts={"Tarif d'achat":"#,##0.00€","Stock":"#,##0","Unités Vendues (Période)":"#,##0","Ventes Moy Hebdo (Période)":"#,##0.00","Ventes Moy Mensuel (Période)":"#,##0.00","Semaines Stock (WoS)":"0.0","Rotation Unités (Proxy)":"0.00","Valeur Stock Actuel (€)":"#,##0.00€","COGS (Période)":"#,##0.00€","Rotation Valeur (Proxy)":"0.00"}
                                format_excel_sheet(ws_r,df_e_r,column_formats=rot_col_fmts)
                            out_b_r.seek(0);f_r_exp=f"{f_base_r}_{sup_e_nm_r}_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx"
                            dl_lbl_r=f"📥 Télécharger ({'Filtrée'if not show_all_f_r else'Complète'})"
                            st.download_button(dl_lbl_r,out_b_r,f_r_exp,"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",key="dl_r_b_t2_dl")
                        except Exception as e_wrt_r:logging.exception(f"Err ExcelWriter rot: {e_wrt_r}");st.error("Erreur export rotation.")
                    else:st.info("Aucune donnée à exporter.")
            else:st.info("Résultats analyse invalidés. Relancer.")

    with tab3: # Vérification Stock Négatif
        st.header("Vérification des Stocks Négatifs")
        st.caption("Analyse tous articles du 'Tableau final'.")
        df_full_neg=st.session_state.get('df_full',None)
        if df_full_neg is None or not isinstance(df_full_neg,pd.DataFrame):st.warning("Données non chargées.")
        elif df_full_neg.empty:st.info("'Tableau final' vide.")
        else:
            stock_c_neg="Stock"
            if stock_c_neg not in df_full_neg.columns:st.error(f"Colonne '{stock_c_neg}' non trouvée.")
            else:
                df_neg_res=df_full_neg[df_full_neg[stock_c_neg]<0].copy()
                if df_neg_res.empty:st.success("✅ Aucun stock négatif.")
                else:
                    st.warning(f"⚠️ **{len(df_neg_res)} article(s) avec stock négatif !**")
                    cols_neg_show=["Fournisseur","AF_RefFourniss","Référence Article","Désignation Article","Stock"]
                    disp_cols_neg=[c for c in cols_neg_show if c in df_neg_res.columns]
                    if not disp_cols_neg:st.error("Cols manquantes affichage négatifs.")
                    else:st.dataframe(df_neg_res[disp_cols_neg].style.format({"Stock":"{:,.0f}"},na_rep="-").apply(lambda s:['background-color:#FADBD8'if s.name==stock_c_neg and val<0 else''for val in s],axis=0))
                    st.markdown("---");st.markdown("#### Exporter Stocks Négatifs")
                    out_b_neg=io.BytesIO();df_exp_neg=df_neg_res[disp_cols_neg].copy()
                    try:
                        with pd.ExcelWriter(out_b_neg,engine="openpyxl")as w_neg:
                            df_exp_neg.to_excel(w_neg,sheet_name="Stocks_Negatifs",index=False)
                            ws_neg=w_neg.sheets["Stocks_Negatifs"]
                            neg_col_fmts={"Stock":"#,##0"}
                            format_excel_sheet(ws_neg,df_exp_neg,column_formats=neg_col_fmts)
                        out_b_neg.seek(0);f_neg_exp=f"stocks_negatifs_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx"
                        st.download_button("📥 Télécharger Liste Négatifs",out_b_neg,f_neg_exp,"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",key="dl_neg_b_t3_dl")
                    except Exception as e_exp_neg:st.error(f"Err export neg: {e_exp_neg}")

    with tab4: # Simulation Forecast
        st.header("Simulation de Forecast Annuel")
        sel_f_t4=render_supplier_checkboxes("tab4",all_sups_data,default_select_all=True) # Shortened
        df_disp_t4=pd.DataFrame() # Shortened
        if sel_f_t4:
            if not df_base_tabs.empty:
                df_disp_t4=df_base_tabs[df_base_tabs["Fournisseur"].isin(sel_f_t4)].copy()
                st.caption(f"{len(df_disp_t4)} art. / {len(sel_f_t4)} fourn.")
        else:st.info("Sélectionner fournisseur(s).")
        st.markdown("---");st.warning("🚨 **Hypothèse:** Saisonnalité mensuelle approx. sur 52 sem. N-1.")
        if df_disp_t4.empty and sel_f_t4:st.warning("Aucun article pour fournisseur(s) sélectionné(s).")
        elif len(id_sem_cols)<52 and not df_disp_t4.empty:st.warning(f"Données histo. < 52 sem ({len(id_sem_cols)}). Simu N-1 impossible.")
        elif not df_disp_t4.empty:
            st.markdown("#### Paramètres Simulation Forecast")
            all_cal_m=list(calendar.month_name)[1:]
            sel_m_f_ui=st.multiselect("📅 Mois simulation:",all_cal_m,default=st.session_state.forecast_selected_months_ui,key="f_m_sel_ui_t4")
            st.session_state.forecast_selected_months_ui=sel_m_f_ui
            sim_t_opts_f=('Simple Progression','Objectif Montant')
            sim_t_f_ui=st.radio("⚙️ Type Simulation:",sim_t_opts_f,horizontal=True,index=st.session_state.forecast_sim_type_radio_index,key="f_sim_t_ui_t4")
            st.session_state.forecast_sim_type_radio_index=sim_t_opts_f.index(sim_t_f_ui)
            prog_pct_f,obj_mt_f=0.0,0.0
            c1_f,c2_f=st.columns(2); # Shortened
            with c1_f:
                if sim_t_f_ui=='Simple Progression':prog_pct_f=st.number_input("📈 Progression (%)",-100.0,value=st.session_state.forecast_progression_percentage_ui,step=0.5,format="%.1f",key="f_prog_pct_ui_t4");st.session_state.forecast_progression_percentage_ui=prog_pct_f
            with c2_f:
                if sim_t_f_ui=='Objectif Montant':obj_mt_f=st.number_input("🎯 Objectif (€) (mois sel.)",0.0,value=st.session_state.forecast_target_amount_ui,step=1000.0,format="%.2f",key="f_target_amt_ui_t4");st.session_state.forecast_target_amount_ui=obj_mt_f
            if st.button("▶️ Lancer Simulation Forecast",key="run_f_sim_btn_t4"):
                if not sel_m_f_ui:st.error("Sélectionner au moins un mois.")
                else:
                    with st.spinner("Simulation forecast..."):df_f_res,gt_f=calculer_forecast_simulation_v3(df_disp_t4,id_sem_cols,sel_m_f_ui,sim_t_f_ui,prog_pct_f,obj_mt_f)
                    if df_f_res is not None:
                        st.success("✅ Simu forecast OK.");st.session_state.forecast_result_df=df_f_res;st.session_state.forecast_grand_total_amount=gt_f
                        st.session_state.forecast_simulation_params_calculated_for={'suppliers':sel_f_t4,'months':sel_m_f_ui,'type':sim_t_f_ui,'prog_pct':prog_pct_f,'obj_amt':obj_mt_f}
                        st.rerun()
                    else:st.error("❌ Simu forecast échouée.")
            if st.session_state.forecast_result_df is not None:
                curr_p_f_ui={'suppliers':sel_f_t4,'months':sel_m_f_ui,'type':sim_t_f_ui,'prog_pct':st.session_state.forecast_progression_percentage_ui if sim_t_f_ui=='Simple Progression'else 0.0,'obj_amt':st.session_state.forecast_target_amount_ui if sim_t_f_ui=='Objectif Montant'else 0.0}
                if st.session_state.forecast_simulation_params_calculated_for==curr_p_f_ui:
                    st.markdown("---");st.markdown("#### Résultats Simulation Forecast")
                    df_f_disp=st.session_state.forecast_result_df;gt_f_disp=st.session_state.forecast_grand_total_amount
                    if df_f_disp.empty:st.info("Aucun résultat simulation.")
                    else:
                        fmts_f={"Tarif d'achat":"{:,.2f}€","Conditionnement":"{:,.0f}"}
                        for m_disp in sel_m_f_ui:
                            if f"Ventes N-1 {m_disp}"in df_f_disp.columns:fmts_f[f"Ventes N-1 {m_disp}"]="{:,.0f}" # Removed leading space
                            if f"Qté Prév. {m_disp}"in df_f_disp.columns:fmts_f[f"Qté Prév. {m_disp}"]="{:,.0f}" # Removed leading space
                            if f"Montant Prév. {m_disp} (€)"in df_f_disp.columns:fmts_f[f"Montant Prév. {m_disp} (€)"]="{:,.2f}€" # Removed leading space
                        for col_n in["Vts N-1 Tot (Mois Sel.)","Qté Tot Prév (Mois Sel.)","Mnt Tot Prév (€) (Mois Sel.)"]:
                            if col_n in df_f_disp.columns:fmts_f[col_n]="{:,.0f}"if"Qté"in col_n or"Vts"in col_n else"{:,.2f}€"
                        try:st.dataframe(df_f_disp.style.format(fmts_f,na_rep="-",thousands=","))
                        except Exception as e_fmt_f:st.error(f"Err format affichage: {e_fmt_f}");st.dataframe(df_f_disp)
                        st.metric(label="💰 Mnt Total Prévisionnel (€) (mois sel.)",value=f"{gt_f_disp:,.2f} €")
                        st.markdown("#### Export Simulation")
                        out_b_f=io.BytesIO();df_e_f=df_f_disp.copy()
                        try:
                            sim_t_fn=sim_t_f_ui.replace(' ','_').lower()
                            with pd.ExcelWriter(out_b_f,engine="openpyxl")as w_f:
                                df_e_f.to_excel(w_f,sheet_name=sanitize_sheet_name(f"Forecast_{sim_t_fn}"),index=False)
                                ws_f=w_f.sheets[sanitize_sheet_name(f"Forecast_{sim_t_fn}")]
                                fcst_col_fmts={"Tarif d'achat":"#,##0.00€","Conditionnement":"#,##0"}
                                for m_disp in sel_m_f_ui:
                                    if f"Ventes N-1 {m_disp}"in df_e_f.columns:fcst_col_fmts[f"Ventes N-1 {m_disp}"]="#,##0"
                                    if f"Qté Prév. {m_disp}"in df_e_f.columns:fcst_col_fmts[f"Qté Prév. {m_disp}"]="#,##0"
                                    if f"Montant Prév. {m_disp} (€)"in df_e_f.columns:fcst_col_fmts[f"Montant Prév. {m_disp} (€)"]="#,##0.00€"
                                if"Vts N-1 Tot (Mois Sel.)"in df_e_f.columns:fcst_col_fmts["Vts N-1 Tot (Mois Sel.)"]="#,##0"
                                if"Qté Tot Prév (Mois Sel.)"in df_e_f.columns:fcst_col_fmts["Qté Tot Prév (Mois Sel.)"]="#,##0"
                                if"Mnt Tot Prév (€) (Mois Sel.)"in df_e_f.columns:fcst_col_fmts["Mnt Tot Prév (€) (Mois Sel.)"]="#,##0.00€"
                                format_excel_sheet(ws_f,df_e_f,column_formats=fcst_col_fmts)
                            out_b_f.seek(0)
                            sup_e_nm_f='multi'if len(sel_f_t4)>1 else(sanitize_sheet_name(sel_f_t4[0])if sel_f_t4 else'NA')
                            f_f_exp=f"forecast_{sim_t_fn}_{sup_e_nm_f}_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx"
                            st.download_button("📥 Télécharger Simulation",out_b_f,f_f_exp,"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",key="dl_f_b_t4_dl")
                        except Exception as eef_f:st.error(f"Err export forecast: {eef_f}")
                else:st.info("Résultats simulation invalidés. Relancer.")

    with tab5: # Suivi Commandes Fournisseurs
        st.header("📄 Suivi des Commandes Fournisseurs")
        if df_suivi_cmds_all is None or df_suivi_cmds_all.empty:
            st.warning("Aucune donnée de suivi (onglet 'Suivi commandes' vide/manquant ou erreur lecture).")
        else:
            sups_in_suivi_list=[] # Shortened
            if"Fournisseur"in df_suivi_cmds_all.columns:sups_in_suivi_list=sorted(df_suivi_cmds_all["Fournisseur"].astype(str).unique().tolist())
            if not sups_in_suivi_list:st.info("Aucun fournisseur trouvé dans données suivi.")
            else:
                st.markdown("Sélectionnez fournisseurs pour archive de suivi:")
                sel_f_t5_ui=render_supplier_checkboxes("tab5",sups_in_suivi_list,default_select_all=False) # Shortened
                if not sel_f_t5_ui:st.info("Sélectionner fournisseur(s) pour générer archive suivi.")
                else:
                    st.markdown("---");st.markdown(f"**{len(sel_f_t5_ui)} fournisseur(s) sélectionné(s) pour export.**")
                    if st.button("📦 Générer et Télécharger Archive ZIP de Suivi",key="gen_suivi_zip_btn_t5"): # Shortened
                        out_cols_s_exp=["Date Pièce BC","N° de pièce","AF_RefFourniss","Désignation Article","Qté Commandées","Date de livraison prévue"] # Shortened
                        src_cols_need_s=["Date Pièce BC","N° de pièce","AF_RefFourniss","Désignation Article","Qté Commandées","Fournisseur"] # Shortened
                        miss_src_cols_s_c=[c for c in src_cols_need_s if c not in df_suivi_cmds_all.columns] # Shortened
                        if miss_src_cols_s_c:st.error(f"Cols sources manquantes ('Suivi cmds'): {', '.join(miss_src_cols_s_c)}. Export impossible.")
                        else:
                            zip_buf=io.BytesIO();files_added_zip=0 # Shortened
                            try:
                                with zipfile.ZipFile(zip_buf,'w',zipfile.ZIP_DEFLATED)as zipf:
                                    for sup_nm_s_exp in sel_f_t5_ui: # Shortened
                                        df_sup_s_exp_d=df_suivi_cmds_all[df_suivi_cmds_all["Fournisseur"]==sup_nm_s_exp].copy() # Shortened
                                        if df_sup_s_exp_d.empty:logging.info(f"Aucune cmd pour {sup_nm_s_exp}, non ajouté ZIP.");continue
                                        df_exp_fin_s=pd.DataFrame(columns=out_cols_s_exp)
                                        if'Date Pièce BC'in df_sup_s_exp_d:df_exp_fin_s["Date Pièce BC"]=pd.to_datetime(df_sup_s_exp_d["Date Pièce BC"],errors='coerce') # Keep datetime
                                        for col_map in["N° de pièce","AF_RefFourniss","Désignation Article","Qté Commandées"]:
                                            if col_map in df_sup_s_exp_d:df_exp_fin_s[col_map]=df_sup_s_exp_d[col_map]
                                        df_exp_fin_s["Date de livraison prévue"]=""
                                        excel_buf_ind=io.BytesIO() # Shortened
                                        with pd.ExcelWriter(excel_buf_ind,engine="openpyxl",date_format='DD/MM/YYYY',datetime_format='DD/MM/YYYY')as writer_ind: # No specific datetime engine needed here for openpyxl
                                            df_to_w=df_exp_fin_s[out_cols_s_exp].copy() # Shortened
                                            sheet_nm=sanitize_sheet_name(f"Suivi_{sup_nm_s_exp}") # Shortened
                                            df_to_w.to_excel(writer_ind,sheet_name=sheet_nm,index=False)
                                            ws=writer_ind.sheets[sheet_nm]
                                            suivi_col_fmts={"Date Pièce BC":"dd/mm/yyyy","Qté Commandées":"#,##0"} # Shortened
                                            format_excel_sheet(ws,df_to_w,column_formats=suivi_col_fmts)
                                        excel_b=excel_buf_ind.getvalue() # Shortened
                                        file_nm_in_zip=f"Suivi_Commande_{sanitize_sheet_name(sup_nm_s_exp)}_{pd.Timestamp.now():%Y%m%d}.xlsx" # Shortened
                                        zipf.writestr(file_nm_in_zip,excel_b)
                                        files_added_zip+=1
                                if files_added_zip>0:
                                    zip_buf.seek(0)
                                    archive_nm=f"Archive_Suivi_Commandes_{pd.Timestamp.now():%Y%m%d_%H%M}.zip" # Shortened
                                    st.download_button(label=f"📥 Télécharger Archive ZIP ({files_added_zip} fichier(s))",data=zip_buf,file_name=archive_nm,mime="application/zip",key="dl_suivi_zip_btn_t5_dl")
                                    st.success(f"{files_added_zip} fichier(s) inclus dans ZIP.") # Shortened
                                else:st.info("Aucun fichier suivi généré.")
                            except Exception as e_zip:logging.exception(f"Err création ZIP suivi: {e_zip}");st.error(f"Err création ZIP: {e_zip}") # Shortened

elif not uploaded_file:
    st.info("👋 Bienvenue ! Chargez votre fichier Excel principal pour démarrer.")
    if st.button("🔄 Réinitialiser l'Application"):
        for k in list(st.session_state.keys()):del st.session_state[k] # Shortened
        st.rerun()
elif 'df_initial_filtered'in st.session_state and not isinstance(st.session_state.df_initial_filtered,pd.DataFrame):
    st.error("Erreur interne : Données filtrées invalides. Rechargez fichier.") # Shortened
    st.session_state.df_full=None
    if st.button("Réessayer"):st.rerun()
