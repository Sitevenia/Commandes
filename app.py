
import streamlit as st
import pandas as pd
import numpy as np
import io
import logging
import re
import openpyxl # Used by pandas for .xlsx, explicit import for type hints
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment
import calendar
import zipfile
from datetime import timedelta # Added for AI date calculations

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
        # Check if streams are still valid and closable before closing
        if hasattr(sys.stdout, 'close') and not getattr(sys.stdout, 'closed', True):
             try:
                 sys.stdout.close()
             except Exception: # Ignore potential errors on close
                 pass
        sys.stdout = self.old_stdout
        if hasattr(sys.stderr, 'close') and not getattr(sys.stderr, 'closed', True):
            try:
                sys.stderr.close()
            except Exception:
                pass
        sys.stderr = self.old_stderr

# --- Helper Functions (safe_read_excel, format_excel_sheet, etc.) ---
def safe_read_excel(uploaded_file, sheet_name, **kwargs):
    """ Safely reads an Excel sheet, returning None if sheet not found or error occurs. """
    try:
        if isinstance(uploaded_file, io.BytesIO): uploaded_file.seek(0)
        file_name_attr = getattr(uploaded_file, 'name', '')
        engine_to_use = 'openpyxl' if file_name_attr.lower().endswith('.xlsx') else None

        logging.debug(f"Attempting to read sheet: '{sheet_name}' from file '{file_name_attr}' with engine '{engine_to_use}' and kwargs: {kwargs}")
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, engine=engine_to_use, **kwargs)

        if df is None:
            logging.error(f"Pandas read_excel returned None for sheet '{sheet_name}'.")
            return None
        logging.debug(f"Read sheet '{sheet_name}'. DataFrame empty: {df.empty}, Columns: {df.columns.tolist()}, Shape: {df.shape}")

        if df.empty and len(df.columns) == 0 and sheet_name is not None:
             logging.warning(f"Sheet '{sheet_name}' was read but has no columns and no rows. Potentially an empty sheet.")
             return pd.DataFrame()

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
        if "zip file" in str(e).lower() or "BadZipFile" in str(type(e).__name__):
             logging.error(f"Error reading sheet '{sheet_name}': Bad zip file (corrupted .xlsx) - {e}")
             st.error(f"❌ Erreur lors de la lecture de l'onglet '{sheet_name}': Fichier .xlsx potentiellement corrompu (erreur zip). Veuillez vérifier le fichier.")
        else:
            logging.error(f"Unexpected error reading sheet '{sheet_name}': {type(e).__name__} - {e}")
            st.error(f"❌ Erreur inattendue ('{type(e).__name__}') lors de la lecture de l'onglet '{sheet_name}': {e}.")
        return None

def format_excel_sheet(worksheet, df, column_formats={}, freeze_header=True, default_float_format="#,##0.00", default_int_format="#,##0", default_date_format="dd/mm/yyyy"):
    if df is None or df.empty:
        logging.warning("Attempted to format sheet with empty/None DataFrame.")
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
        num_format_to_apply = None
        try:
            header_length = len(str(col_name))
            non_na_series = df[col_name].dropna()
            sample_data = non_na_series.sample(min(len(non_na_series), 20)) if not non_na_series.empty else pd.Series([], dtype='object')

            data_length = 0
            if not sample_data.empty:
                try:
                    data_length = sample_data.astype(str).map(len).max()
                except Exception:
                    data_length = 0

            data_length = data_length if pd.notna(data_length) else 0
            current_max_len = max(header_length, data_length) + 3
            final_width = min(max(current_max_len, 10), 50)
            worksheet.column_dimensions[col_letter].width = final_width
        except Exception as e_width:
            logging.warning(f"Could not automatically set width for column {col_name}: {e_width}")
            worksheet.column_dimensions[col_letter].width = 15

        specific_format = column_formats.get(col_name)
        try:
            col_dtype = df[col_name].dtype
        except KeyError:
            logging.warning(f"Column '{col_name}' not found in DataFrame during formatting.")
            continue

        if specific_format:
            num_format_to_apply = specific_format
        elif pd.api.types.is_integer_dtype(col_dtype):
            num_format_to_apply = default_int_format
        elif pd.api.types.is_float_dtype(col_dtype):
            num_format_to_apply = default_float_format
        elif pd.api.types.is_datetime64_any_dtype(col_dtype) or \
             (not df[col_name].empty and isinstance(df[col_name].dropna().iloc[0] if not df[col_name].dropna().empty else None, pd.Timestamp)):
            num_format_to_apply = default_date_format

        for row_idx in range(2, worksheet.max_row + 1):
            cell = worksheet[f"{col_letter}{row_idx}"]
            cell.alignment = data_alignment
            if num_format_to_apply and cell.value is not None and not str(cell.value).startswith('='):
                try:
                    cell.number_format = num_format_to_apply
                except Exception as e_format_cell:
                    logging.warning(f"Could not apply format '{num_format_to_apply}' to cell {col_letter}{row_idx} for column {col_name}: {e_format_cell}")

    if freeze_header:
        worksheet.freeze_panes = worksheet['A2']

def calculer_quantite_a_commander(df, semaine_columns, montant_minimum_input, duree_semaines):
    try:
        if not isinstance(df, pd.DataFrame) or df.empty:
            st.info("Aucune donnée fournie pour le calcul des quantités.")
            return None

        required_cols = ["Stock", "Conditionnement", "Tarif d'achat"] + semaine_columns
        missing_cols = [c for c in required_cols if c not in df.columns]
        if missing_cols:
            st.error(f"Colonnes manquantes pour le calcul: {', '.join(missing_cols)}")
            return None

        if not semaine_columns:
            st.error("Aucune colonne 'semaine' identifiée pour le calcul.")
            return None

        df_calc = df.copy()
        for col in required_cols:
            df_calc[col] = pd.to_numeric(df_calc[col], errors='coerce').replace([np.inf, -np.inf], np.nan).fillna(0)

        num_semaines_total = len(semaine_columns)
        ventes_N1_total_series = df_calc[semaine_columns].sum(axis=1)

        if num_semaines_total >= 64:
            ventes_12_N1_similaires = df_calc[semaine_columns[-64:-52]].sum(axis=1)
            ventes_12_N1_decalees = df_calc[semaine_columns[-52:-40]].sum(axis=1)
            moy_12_N1_similaires = ventes_12_N1_similaires / 12
            moy_12_N1_decalees = ventes_12_N1_decalees / 12
        else:
            ventes_12_N1_similaires, ventes_12_N1_decalees, moy_12_N1_similaires, moy_12_N1_decalees = (pd.Series(0.0, index=df_calc.index) for _ in range(4))
            if num_semaines_total > 0 :
                 logging.info(f"Moins de 64 semaines de données ({num_semaines_total}), les moyennes N-1 ne seront pas précises ou nulles.")

        nb_semaines_recentes = min(num_semaines_total, 12)
        if nb_semaines_recentes > 0:
            ventes_12_dernieres = df_calc[semaine_columns[-nb_semaines_recentes:]].sum(axis=1)
            moy_12_dernieres = ventes_12_dernieres / nb_semaines_recentes
        else:
            ventes_12_dernieres, moy_12_dernieres = (pd.Series(0.0, index=df_calc.index) for _ in range(2))

        qte_hebdo_ponderee = (0.5 * moy_12_dernieres + 0.2 * moy_12_N1_similaires + 0.3 * moy_12_N1_decalees)
        qte_necessaire_periode = qte_hebdo_ponderee * duree_semaines
        qte_a_commander_suggest = (qte_necessaire_periode - df_calc["Stock"]).apply(lambda x: max(0, x))

        qte_commandee_final_list = []
        conditionnement_series = df_calc["Conditionnement"]
        stock_series = df_calc["Stock"]
        tarif_series = df_calc["Tarif d'achat"]

        for i in range(len(qte_a_commander_suggest)):
            cond = conditionnement_series.iloc[i]
            q_sugg = qte_a_commander_suggest.iloc[i]
            q_final_item = 0
            if q_sugg > 0:
                if cond > 0:
                    q_final_item = int(np.ceil(q_sugg / cond) * cond)
                else:
                    ref_art = df_calc.get('Référence Article', pd.Series(['N/A'], index=df_calc.index)).iloc[i]
                    logging.warning(f"Article index {df_calc.index[i]} (Ref: {ref_art}) "
                                    f"Quantité suggérée {q_sugg:.2f} ignorée car conditionnement={cond}.")
            qte_commandee_final_list.append(q_final_item)

        qte_commandee_final_series = pd.Series(qte_commandee_final_list, index=df_calc.index)

        if nb_semaines_recentes > 0:
            for i in range(len(qte_commandee_final_series)):
                cond = conditionnement_series.iloc[i]
                if cond > 0 and stock_series.iloc[i] <= 1:
                    ventes_recentes_item_non_nulles = (df_calc[semaine_columns[-nb_semaines_recentes:]].iloc[i] > 0).sum()
                    if ventes_recentes_item_non_nulles >= 2:
                        qte_commandee_final_series.iloc[i] = max(qte_commandee_final_series.iloc[i], cond)

        for i in range(len(qte_commandee_final_series)):
            if ventes_N1_total_series.iloc[i] < 6 and ventes_12_dernieres.iloc[i] < 2:
                qte_commandee_final_series.iloc[i] = 0

        montant_actuel_commande = (qte_commandee_final_series * tarif_series).sum()

        if montant_minimum_input > 0 and montant_actuel_commande < montant_minimum_input:
            articles_eligibles_pour_increment = []
            for i in range(len(qte_commandee_final_series)):
                if qte_commandee_final_series.iloc[i] > 0 and conditionnement_series.iloc[i] > 0 and tarif_series.iloc[i] > 0:
                    articles_eligibles_pour_increment.append(i) # Store list index

            if not articles_eligibles_pour_increment:
                if montant_actuel_commande < montant_minimum_input:
                    st.warning(f"Impossible d'atteindre le minimum de commande de {montant_minimum_input:,.2f}€. "
                               f"Montant actuel: {montant_actuel_commande:,.2f}€. Aucun article éligible pour incrémentation.")
            else:
                idx_ptr_eligible = 0
                max_iterations_loop = len(articles_eligibles_pour_increment) * 20 + 1
                iterations_count = 0
                qte_commandee_temp_list_adj = qte_commandee_final_series.tolist()

                while montant_actuel_commande < montant_minimum_input and iterations_count < max_iterations_loop:
                    iterations_count += 1
                    list_index_item_to_inc = articles_eligibles_pour_increment[idx_ptr_eligible]

                    cond_item = conditionnement_series.iloc[list_index_item_to_inc]
                    tarif_item = tarif_series.iloc[list_index_item_to_inc]

                    if cond_item > 0 and tarif_item > 0 :
                        qte_commandee_temp_list_adj[list_index_item_to_inc] += cond_item
                        montant_actuel_commande += (cond_item * tarif_item)
                    else:
                        logging.warning(f"Skipping increment for item index {list_index_item_to_inc} due to invalid cond/price.")

                    idx_ptr_eligible = (idx_ptr_eligible + 1) % len(articles_eligibles_pour_increment)

                qte_commandee_final_series = pd.Series(qte_commandee_temp_list_adj, index=df_calc.index)

                if iterations_count >= max_iterations_loop and montant_actuel_commande < montant_minimum_input:
                    st.error(f"Ajustement pour minimum: Nombre maximum d'itérations ({max_iterations_loop}) atteint. "
                             f"Montant actuel: {montant_actuel_commande:,.2f}€ / Requis: {montant_minimum_input:,.2f}€.")

        montant_final_commande = (qte_commandee_final_series * tarif_series).sum()
        return (qte_commandee_final_series, ventes_N1_total_series, ventes_12_N1_similaires, ventes_12_dernieres, montant_final_commande)

    except KeyError as e:
        st.error(f"Erreur de clé lors du calcul des quantités: '{e}'. Vérifiez les noms de colonnes.")
        logging.exception(f"KeyError in calculer_quantite_a_commander: {e}")
        return None
    except Exception as e:
        st.error(f"Erreur inattendue lors du calcul des quantités: {type(e).__name__} - {e}")
        logging.exception("Exception in calculer_quantite_a_commander:")
        return None

def calculer_rotation_stock(df, semaine_columns, periode_semaines_analyse):
    try:
        if not isinstance(df, pd.DataFrame) or df.empty:
            st.info("Aucune donnée fournie pour l'analyse de rotation.")
            return pd.DataFrame()

        required_cols = ["Stock", "Tarif d'achat"]
        missing_cols = [c for c in required_cols if c not in df.columns]
        if missing_cols:
            st.error(f"Colonnes manquantes pour l'analyse de rotation: {', '.join(missing_cols)}")
            return None

        df_rotation = df.copy()
        semaines_pour_analyse, nb_semaines_analyse_effectif = [], 0
        if periode_semaines_analyse and periode_semaines_analyse > 0:
            if len(semaine_columns) >= periode_semaines_analyse:
                semaines_pour_analyse = semaine_columns[-periode_semaines_analyse:]
                nb_semaines_analyse_effectif = periode_semaines_analyse
            else:
                semaines_pour_analyse = semaine_columns
                nb_semaines_analyse_effectif = len(semaine_columns)
                st.caption(f"Période d'analyse ajustée à {nb_semaines_analyse_effectif} semaines (toutes les données disponibles).")
        else:
            semaines_pour_analyse = semaine_columns
            nb_semaines_analyse_effectif = len(semaine_columns)

        if not semaines_pour_analyse:
            st.warning("Aucune colonne de vente disponible pour l'analyse de rotation.")
            metric_cols_definition = ["Unités Vendues (Période)", "Ventes Moy Hebdo (Période)", "Ventes Moy Mensuel (Période)", "Semaines Stock (WoS)", "Rotation Unités (Proxy)", "COGS (Période)", "Valeur Stock Actuel (€)", "Rotation Valeur (Proxy)"]
            for m_col_name in metric_cols_definition: df_rotation[m_col_name] = 0.0
            return df_rotation

        for col_name in semaines_pour_analyse:
             if col_name in df_rotation.columns:
                 df_rotation[col_name] = pd.to_numeric(df_rotation[col_name], errors='coerce').fillna(0)
             else:
                 logging.warning(f"Sales column '{col_name}' expected but not found in rotation DataFrame.")
                 df_rotation[col_name] = 0.0
        df_rotation["Stock"] = pd.to_numeric(df_rotation["Stock"], errors='coerce').fillna(0)
        df_rotation["Tarif d'achat"] = pd.to_numeric(df_rotation["Tarif d'achat"], errors='coerce').fillna(0)

        valid_sales_cols_in_df = [col for col in semaines_pour_analyse if col in df_rotation.columns]
        df_rotation["Unités Vendues (Période)"] = df_rotation[valid_sales_cols_in_df].sum(axis=1) if valid_sales_cols_in_df else 0.0

        df_rotation["Ventes Moy Hebdo (Période)"] = df_rotation["Unités Vendues (Période)"] / nb_semaines_analyse_effectif if nb_semaines_analyse_effectif > 0 else 0.0
        df_rotation["Ventes Moy Mensuel (Période)"] = df_rotation["Ventes Moy Hebdo (Période)"] * (52.0 / 12.0)

        denominator_wos = df_rotation["Ventes Moy Hebdo (Période)"]
        current_stock_rot = df_rotation["Stock"]
        df_rotation["Semaines Stock (WoS)"] = np.divide(
            current_stock_rot,
            denominator_wos,
            out=np.full_like(current_stock_rot, np.inf, dtype=np.float64),
            where=denominator_wos != 0
        )
        df_rotation.loc[current_stock_rot <= 0, "Semaines Stock (WoS)"] = 0.0

        denominator_rot_units = current_stock_rot
        df_rotation["Rotation Unités (Proxy)"] = np.divide(
            df_rotation["Unités Vendues (Période)"],
            denominator_rot_units,
            out=np.full_like(denominator_rot_units, np.inf, dtype=np.float64),
            where=denominator_rot_units != 0
        )
        df_rotation.loc[(df_rotation["Unités Vendues (Période)"] <= 0) & (denominator_rot_units <= 0), "Rotation Unités (Proxy)"] = 0.0
        df_rotation.loc[(df_rotation["Unités Vendues (Période)"] <= 0) & (denominator_rot_units > 0), "Rotation Unités (Proxy)"] = 0.0

        df_rotation["COGS (Période)"] = df_rotation["Unités Vendues (Période)"] * df_rotation["Tarif d'achat"]
        df_rotation["Valeur Stock Actuel (€)"] = current_stock_rot * df_rotation["Tarif d'achat"]

        denominator_rot_value = df_rotation["Valeur Stock Actuel (€)"]
        df_rotation["Rotation Valeur (Proxy)"] = np.divide(
            df_rotation["COGS (Période)"],
            denominator_rot_value,
            out=np.full_like(denominator_rot_value, np.inf, dtype=np.float64),
            where=denominator_rot_value != 0
        )
        df_rotation.loc[(df_rotation["COGS (Période)"] <= 0) & (denominator_rot_value <= 0), "Rotation Valeur (Proxy)"] = 0.0
        df_rotation.loc[(df_rotation["COGS (Période)"] <= 0) & (denominator_rot_value > 0), "Rotation Valeur (Proxy)"] = 0.0

        return df_rotation

    except KeyError as e:
        st.error(f"Erreur de clé lors de l'analyse de rotation: '{e}'.")
        logging.exception(f"KeyError in calculer_rotation_stock: {e}")
        return None
    except Exception as e:
        st.error(f"Erreur inattendue lors de l'analyse de rotation: {type(e).__name__} - {e}")
        logging.exception("Error in calculer_rotation_stock:")
        return None

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
        if not isinstance(df, pd.DataFrame) or df.empty:
            st.warning("Aucune donnée produit pour la simulation de forecast.")
            return None, 0.0

        if not all_historical_semaine_columns:
             st.error("Aucune colonne de ventes historiques fournie pour la simulation.")
             return None, 0.0

        if not selected_months_list:
            st.warning("Veuillez sélectionner au moins un mois pour la simulation.")
            return None, 0.0

        required_data_cols = ["Référence Article", "Désignation Article", "Conditionnement", "Tarif d'achat", "Fournisseur"]
        missing_data_cols = [c for c in required_data_cols if c not in df.columns]
        if missing_data_cols:
            st.error(f"Colonnes de données manquantes pour la simulation: {', '.join(missing_data_cols)}")
            return None, 0.0

        parsed_week_col_objects = []
        available_years = set()
        for col_name_str in all_historical_semaine_columns:
            if isinstance(col_name_str, str):
                match = re.match(r"(\d{4})[SW]?(\d{1,2})", col_name_str, re.IGNORECASE)
                if match:
                    year, week_num = int(match.group(1)), int(match.group(2))
                    if 1 <= week_num <= 53:
                        available_years.add(year)
                        parsed_week_col_objects.append({'year': year, 'week': week_num, 'col': col_name_str, 'sort_key': year * 100 + week_num})

        if not available_years:
            st.error("Impossible de déterminer les années. Format: 'YYYYWW' ou 'YYYYSwW'.")
            return None, 0.0

        parsed_week_col_objects.sort(key=lambda x: x['sort_key'])
        current_year_n = max(available_years) if available_years else 0
        previous_year_n_minus_1 = current_year_n - 1
        st.caption(f"Simulation basée sur N-1 (N: {current_year_n}, N-1: {previous_year_n_minus_1})")

        n1_week_data_objects = [item for item in parsed_week_col_objects if item['year'] == previous_year_n_minus_1]
        if len(n1_week_data_objects) < 52:
            st.error(f"Données N-1 ({previous_year_n_minus_1}) < 52 sem. ({len(n1_week_data_objects)}).")
            return None, 0.0

        n1_week_column_names_for_mapping = [item['col'] for item in n1_week_data_objects[:52]]
        df_simulation_results = df[required_data_cols].copy()
        df_simulation_results["Tarif d'achat"] = pd.to_numeric(df_simulation_results["Tarif d'achat"], errors='coerce').fillna(0)
        df_simulation_results["Conditionnement"] = pd.to_numeric(df_simulation_results["Conditionnement"], errors='coerce').fillna(1).apply(lambda x: 1 if x <= 0 else int(x))

        if not all(c in df.columns for c in n1_week_column_names_for_mapping):
            missing_df_cols = [c for c in n1_week_column_names_for_mapping if c not in df.columns]
            st.error(f"Erreur interne: Cols N-1 mappées non trouvées dans données de base: {', '.join(missing_df_cols)}")
            return None, 0.0

        df_n1_sales_only = df[n1_week_column_names_for_mapping].copy()
        for col_n1 in n1_week_column_names_for_mapping:
            if col_n1 in df_n1_sales_only.columns:
                df_n1_sales_only[col_n1] = pd.to_numeric(df_n1_sales_only[col_n1], errors='coerce').fillna(0)
            else:
                df_n1_sales_only[col_n1] = 0.0

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
            if objectif_montant_val <= 0:
                st.error("Objectif Montant > 0 requis.")
                return None, 0.0

            total_n1_value_all_selected_months = (total_n1_sales_for_selected_months_series * df_simulation_results["Tarif d'achat"]).sum()
            if total_n1_value_all_selected_months <= 0:
                st.warning("Ventes N-1 (valeur) nulles pour mois sel. Répartition montant objectif basée sur prix et nombre de mois.")
                num_selected_m = len(selected_months_list)
                if num_selected_m == 0: return None, 0.0
                num_items_gt_zero_price = (df_simulation_results["Tarif d'achat"] > 0).sum()
                if num_items_gt_zero_price == 0:
                    st.warning("Aucun article avec prix > 0. Impossible de répartir l'objectif montant.")
                    target_amount_per_month_item_avg = 0.0
                else:
                     target_amount_per_month_item_avg = objectif_montant_val / num_selected_m / num_items_gt_zero_price

                for m_name_fcst in selected_months_list:
                    base_monthly_forecast_qty_map[m_name_fcst] = np.divide(
                        target_amount_per_month_item_avg,
                        df_simulation_results["Tarif d'achat"],
                        out=np.zeros_like(df_simulation_results["Tarif d'achat"], dtype=float),
                        where=df_simulation_results["Tarif d'achat"] != 0)
            else:
                for m_name_fcst in selected_months_list:
                    monthly_n1_value_series = (monthly_n1_sales_map_for_selected_months.get(m_name_fcst, pd.Series(0.0, index=df_simulation_results.index)) * df_simulation_results["Tarif d'achat"])
                    month_value_contribution_factor = np.divide(monthly_n1_value_series.sum(), total_n1_value_all_selected_months, out=np.array([0.0]), where=total_n1_value_all_selected_months !=0)[0]
                    target_amount_this_month_global = objectif_montant_val * month_value_contribution_factor
                    item_contribution_in_month_value_factor = np.divide(monthly_n1_value_series, monthly_n1_value_series.sum(), out=np.zeros_like(monthly_n1_value_series,dtype=float), where=monthly_n1_value_series.sum() !=0)
                    target_amount_per_item_this_month = target_amount_this_month_global * item_contribution_in_month_value_factor
                    base_monthly_forecast_qty_map[m_name_fcst] = np.divide(target_amount_per_item_this_month, df_simulation_results["Tarif d'achat"], out=np.zeros_like(df_simulation_results["Tarif d'achat"], dtype=float), where=df_simulation_results["Tarif d'achat"] != 0)
        else:
            st.error(f"Type simu non reconnu: '{sim_type_str}'.")
            return None, 0.0

        total_adjusted_qty_all_months_series = pd.Series(0.0, index=df_simulation_results.index)
        total_final_amount_all_months_series = pd.Series(0.0, index=df_simulation_results.index)
        for m_name_fcst in selected_months_list:
            forecast_qty_col_name, forecast_amt_col_name = f"Qté Prév. {m_name_fcst}", f"Montant Prév. {m_name_fcst} (€)"
            base_qty_series = base_monthly_forecast_qty_map.get(m_name_fcst, pd.Series(0.0, index=df_simulation_results.index))
            base_qty_series = pd.to_numeric(base_qty_series, errors='coerce').fillna(0)
            conditionnement_series_sim = df_simulation_results["Conditionnement"]
            adjusted_qty_series = (np.ceil(np.divide(base_qty_series, conditionnement_series_sim, out=np.zeros_like(base_qty_series, dtype=float), where=conditionnement_series_sim != 0)) * conditionnement_series_sim).fillna(0).astype(int)
            df_simulation_results[forecast_qty_col_name] = adjusted_qty_series
            df_simulation_results[forecast_amt_col_name] = adjusted_qty_series * df_simulation_results["Tarif d'achat"]
            total_adjusted_qty_all_months_series += adjusted_qty_series
            total_final_amount_all_months_series += df_simulation_results[forecast_amt_col_name]

        df_simulation_results["Qté Totale Prév. (Mois Sel.)"] = total_adjusted_qty_all_months_series
        df_simulation_results["Montant Total Prév. (€) (Mois Sel.)"] = total_final_amount_all_months_series

        id_cols_display = ["Fournisseur", "Référence Article", "Désignation Article", "Conditionnement", "Tarif d'achat"]
        n1_sales_cols_display = sorted([f"Ventes N-1 {m}" for m in selected_months_list if f"Ventes N-1 {m}" in df_simulation_results.columns])
        qty_forecast_cols_display = sorted([f"Qté Prév. {m}" for m in selected_months_list if f"Qté Prév. {m}" in df_simulation_results.columns])
        amt_forecast_cols_display = sorted([f"Montant Prév. {m} (€)" for m in selected_months_list if f"Montant Prév. {m} (€)" in df_simulation_results.columns])

        df_simulation_results.rename(columns={"Qté Totale Prév. (Mois Sel.)": "Qté Tot Prév (Mois Sel.)", "Montant Total Prév. (€) (Mois Sel.)": "Mnt Tot Prév (€) (Mois Sel.)"}, inplace=True)
        total_summary_cols_display = ["Vts N-1 Tot (Mois Sel.)", "Qté Tot Prév (Mois Sel.)", "Mnt Tot Prév (€) (Mois Sel.)"]

        final_ordered_columns = id_cols_display + total_summary_cols_display + n1_sales_cols_display + qty_forecast_cols_display + amt_forecast_cols_display
        final_ordered_columns_existing = [c for c in final_ordered_columns if c in df_simulation_results.columns]
        grand_total_forecast_amount = total_final_amount_all_months_series.sum()

        return df_simulation_results[final_ordered_columns_existing], grand_total_forecast_amount

    except KeyError as e:
        st.error(f"Err clé (simu fcst): '{e}'.")
        logging.exception(f"KeyError in calc_fcst_sim_v3: {e}")
        return None, 0.0
    except Exception as e:
        st.error(f"Err inattendue (simu fcst): {type(e).__name__} - {e}")
        logging.exception("Error in calc_fcst_sim_v3:")
        return None, 0.0

def sanitize_sheet_name(name):
    if not isinstance(name, str): name = str(name)
    s_name = re.sub(r'[\[\]:*?/\\<>|"]', '_', name)
    if s_name.startswith("'"): s_name = "_" + s_name[1:]
    if s_name.endswith("'"): s_name = s_name[:-1] + "_"
    return s_name[:31]

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

def sanitize_supplier_key(supplier_name_str):
    if not isinstance(supplier_name_str, str): supplier_name_str = str(supplier_name_str)
    s_key = re.sub(r'\W+', '_', supplier_name_str)
    s_key = re.sub(r'^_+|_+$', '', s_key)
    s_key = re.sub(r'_+', '_', s_key)
    return s_key if s_key else "invalid_supplier_key_name"

# --- AI Helper Functions ---
def parse_week_column_to_date(col_name_str):
    if not isinstance(col_name_str, str): col_name_str = str(col_name_str)
    match_sw = re.match(r"(\d{4})[SW](\d{1,2})", col_name_str, re.IGNORECASE)
    match_plain = re.match(r"(\d{4})(\d{2})", col_name_str)
    year, week_num = None, None
    if match_sw: year, week_num = int(match_sw.group(1)), int(match_sw.group(2))
    elif match_plain:
        potential_year, potential_week = int(match_plain.group(1)), int(match_plain.group(2))
        if 1 <= potential_week <= 53 and 1900 < potential_year < 2200 :
            year, week_num = potential_year, potential_week
        else: return None
    else: return None
    if year and week_num and (1 <= week_num <= 53):
        try:
            date_str_iso = f"{year}-W{week_num:02}-1"
            return pd.to_datetime(date_str_iso, format="%G-W%V-%u")
        except ValueError as e:
            logging.error(f"Err converting {year}W{week_num} (from '{col_name_str}') to date: {e}")
            return None
    return None

def ai_calculate_order_quantities(df_products_for_ai, historical_semaine_cols, num_forecast_weeks,
                                  min_order_amount_for_subset=0.0, apply_special_rules=True):
    if not PROPHET_AVAILABLE:
        st.error("Librairie Prophet (IA) non installée. Prévision IA désactivée.")
        return None, 0.0
    if df_products_for_ai.empty:
        st.info("Aucune donnée produit pour prévision IA.")
        return None, 0.0

    base_req_cols = ["Stock", "Conditionnement", "Tarif d'achat", "Référence Article"]
    missing_base = [c for c in base_req_cols if c not in df_products_for_ai.columns and c != "Référence Article"]
    if missing_base:
        st.error(f"Cols de base manquantes (calcul IA): {', '.join(missing_base)}")
        return None, 0.0

    df_calc_ai = df_products_for_ai.copy()
    for col_op in ["Stock", "Conditionnement", "Tarif d'achat"]:
        if col_op in df_calc_ai.columns:
             df_calc_ai[col_op] = pd.to_numeric(df_calc_ai[col_op], errors='coerce').fillna(0)
        else:
             st.error(f"Colonne critique '{col_op}' manquante pour le calcul IA.")
             return None, 0.0
    df_calc_ai["Conditionnement"] = df_calc_ai["Conditionnement"].apply(lambda x: int(x) if x > 0 else 1)

    parsed_sales_dates = []
    valid_sales_cols_for_model = []
    for col_hist in historical_semaine_cols:
        parsed_dt_obj = parse_week_column_to_date(col_hist)
        if parsed_dt_obj:
            parsed_sales_dates.append({'date': parsed_dt_obj, 'col_name': col_hist})
            valid_sales_cols_for_model.append(col_hist)
        else: logging.warning(f"Colonne '{col_hist}' ignorée pour IA (parsing date échoué).")

    if not parsed_sales_dates:
        st.error("Aucune colonne de ventes historiques n'a pu être interprétée comme date pour l'IA.")
        return None, 0.0
    parsed_sales_df_map = pd.DataFrame(parsed_sales_dates).sort_values(by='date').reset_index(drop=True)

    for col_valid_ts in valid_sales_cols_for_model:
        if col_valid_ts in df_calc_ai.columns:
            df_calc_ai[col_valid_ts] = pd.to_numeric(df_calc_ai[col_valid_ts], errors='coerce')
        else:
             logging.warning(f"Historical sales column '{col_valid_ts}' not found in input DataFrame for AI calculation.")
             df_calc_ai[col_valid_ts] = np.nan

    df_calc_ai["Qté Cmdée (IA)"] = 0
    df_calc_ai["Forecast Ventes (IA)"] = 0.0

    num_prods = len(df_calc_ai)
    # Note: Progress bar update within the loop might slow down significantly for large datasets.
    # Consider updating less frequently if performance is an issue.
    progress_bar_placeholder = st.empty()

    for i, (prod_idx, prod_row) in enumerate(df_calc_ai.iterrows()):
        # Update progress bar less frequently if needed
        # if i % 10 == 0 or i == num_prods - 1: # Update every 10 items or on the last item
        progress_bar_placeholder.progress((i + 1) / num_prods, text=f"Prévision IA: Article {i+1}/{num_prods}")

        prod_ref_log = prod_row.get("Référence Article", f"Index {prod_idx}")
        logging.info(f"Prévision IA pour: {prod_ref_log}")

        prod_ts_hist = [{'ds': ps_row['date'], 'y': prod_row.get(ps_row['col_name'], np.nan)} for _, ps_row in parsed_sales_df_map.iterrows()]
        prod_ts_df_fit = pd.DataFrame(prod_ts_hist).dropna(subset=['ds'])

        if prod_ts_df_fit['y'].notna().sum() < 12:
            logging.warning(f"Produit {prod_ref_log}: <12 points de ventes non-nulles. Prévision IA ignorée.")
            df_calc_ai.loc[prod_idx, "Qté Cmdée (IA)"] = 0; df_calc_ai.loc[prod_idx, "Forecast Ventes (IA)"] = 0.0
            continue
        try:
            model_prophet = Prophet(uncertainty_samples=0)
            if not prod_ts_df_fit.empty and (prod_ts_df_fit['ds'].max() - prod_ts_df_fit['ds'].min()) >= pd.Timedelta(days=365 + 180):
                model_prophet.add_seasonality(name='yearly', period=365.25, fourier_order=10)

            with SuppressStdoutStderr(): model_prophet.fit(prod_ts_df_fit[['ds', 'y']].dropna(subset=['y']))

            future_df = model_prophet.make_future_dataframe(periods=num_forecast_weeks, freq='W-MON')
            forecast_df_res = model_prophet.predict(future_df)
            total_fcst_period = forecast_df_res['yhat'].iloc[-num_forecast_weeks:].sum()
            total_fcst_period = max(0, total_fcst_period)
            df_calc_ai.loc[prod_idx, "Forecast Ventes (IA)"] = total_fcst_period

            stock_item = prod_row["Stock"]; package_item = prod_row["Conditionnement"]
            needed_raw = total_fcst_period - stock_item
            order_qty_item_ia = 0
            if needed_raw > 0:
                if package_item > 0: order_qty_item_ia = int(np.ceil(needed_raw / package_item) * package_item)
                else: logging.warning(f"Produit {prod_ref_log}: Cond. {package_item} invalide. Cmd IA=0.")

            if apply_special_rules and order_qty_item_ia == 0 and stock_item <= 1 and package_item > 0:
                recent_sales_cols_chk = [psc_row['col_name'] for psc_row in parsed_sales_df_map.tail(12).to_dict('records')]
                actual_recent_cols = [c for c in recent_sales_cols_chk if c in df_calc_ai.columns]
                if actual_recent_cols and df_calc_ai.loc[prod_idx, actual_recent_cols].sum() > 0:
                    order_qty_item_ia = package_item
                    logging.info(f"Produit {prod_ref_log}: Stock bas, vts récentes, fcst IA=0. Forçage à 1 cond ({package_item}).")
            df_calc_ai.loc[prod_idx, "Qté Cmdée (IA)"] = order_qty_item_ia
        except Exception as e_ph:
            logging.error(f"Erreur Prophet pour {prod_ref_log}: {e_ph}")
            df_calc_ai.loc[prod_idx, "Qté Cmdée (IA)"] = 0; df_calc_ai.loc[prod_idx, "Forecast Ventes (IA)"] = 0.0

    progress_bar_placeholder.empty()

    df_calc_ai["Total Cmd (€) (IA)"] = df_calc_ai["Qté Cmdée (IA)"] * df_calc_ai["Tarif d'achat"]
    current_total_amount_ia = df_calc_ai["Total Cmd (€) (IA)"].sum()

    if min_order_amount_for_subset > 0 and current_total_amount_ia < min_order_amount_for_subset:
        logging.info(f"Ajustement IA pour min cmd: {min_order_amount_for_subset:,.2f}€. Actuel: {current_total_amount_ia:,.2f}€")
        eligible_inc_indices = df_calc_ai[(df_calc_ai["Qté Cmdée (IA)"] > 0) & (df_calc_ai["Conditionnement"] > 0) & (df_calc_ai["Tarif d'achat"] > 0)].index.tolist()
        if not eligible_inc_indices:
            st.warning(f"Min cmd (IA) de {min_order_amount_for_subset:,.2f}€ non atteint. Aucun article éligible pour incrément.")
        else:
            item_ptr_adj = 0; max_adj_iter = len(eligible_inc_indices) * 20 + 1; current_adj_iter = 0
            qtes_cmdees_ia_series_adj = df_calc_ai["Qté Cmdée (IA)"].copy()

            while current_total_amount_ia < min_order_amount_for_subset and current_adj_iter < max_adj_iter:
                current_adj_iter += 1
                df_item_idx_inc = eligible_inc_indices[item_ptr_adj]
                pkg_adj = df_calc_ai.loc[df_item_idx_inc, "Conditionnement"]
                price_adj = df_calc_ai.loc[df_item_idx_inc, "Tarif d'achat"]

                if pkg_adj > 0 and price_adj > 0:
                    qtes_cmdees_ia_series_adj.loc[df_item_idx_inc] += pkg_adj
                    current_total_amount_ia += (pkg_adj * price_adj)
                else:
                    logging.warning(f"Skipping min order increment for item index {df_item_idx_inc} due to invalid pkg/price.")

                item_ptr_adj = (item_ptr_adj + 1) % len(eligible_inc_indices)

            df_calc_ai["Qté Cmdée (IA)"] = qtes_cmdees_ia_series_adj

            if current_adj_iter >= max_adj_iter and current_total_amount_ia < min_order_amount_for_subset:
                 st.error(f"Ajustement min (IA): Max itérations atteintes. Actuel: {current_total_amount_ia:,.2f}€ / Requis: {min_order_amount_for_subset:,.2f}€.")
            else: logging.info(f"Montant après ajustement IA pour min: {current_total_amount_ia:,.2f}€")

            df_calc_ai["Total Cmd (€) (IA)"] = df_calc_ai["Qté Cmdée (IA)"] * df_calc_ai["Tarif d'achat"]
            current_total_amount_ia = df_calc_ai["Total Cmd (€) (IA)"].sum()

    df_calc_ai["Stock Terme (IA)"] = df_calc_ai["Stock"] + df_calc_ai["Qté Cmdée (IA)"]
    return df_calc_ai, current_total_amount_ia

# --- Streamlit App ---
st.set_page_config(page_title="Forecast & Rotation App", layout="wide")
st.title("📦 Application Prévision Commande, Analyse Rotation & Suivi")
uploaded_file = st.file_uploader("📁 Charger le fichier Excel principal", type=["xlsx", "xls"], key="main_file_uploader")

def get_default_session_state():
    return {
        'df_full': None, 'min_order_dict': {}, 'df_initial_filtered': pd.DataFrame(),
        'all_available_semaine_columns': [], 'unique_suppliers_list': [],
        'commande_result_df': None, 'commande_calculated_total_amount': 0.0,
        'commande_suppliers_calculated_for': [], 'commande_params_calculated_for': {},
        'ai_commande_result_df': None, 'ai_commande_total_amount': 0.0,
        'ai_commande_params_calculated_for': {}, 'ai_forecast_weeks_val': 4, 'ai_min_order_val': 0.0,
        'ai_stock_reduc_target_val': 0.0,
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
        if df_full_read is None or df_full_read.empty:
            st.error("❌ Échec lecture 'Tableau final' ou onglet vide.")
            st.stop()

        req_tf_cols_check = ["Stock", "Fournisseur", "AF_RefFourniss", "Tarif d'achat", "Conditionnement", "Référence Article", "Désignation Article"]
        missing_tf_check = [c for c in req_tf_cols_check if c not in df_full_read.columns]
        if missing_tf_check:
            st.error(f"❌ Cols manquantes ('TF'): {', '.join(missing_tf_check)}. Vérifiez ligne en-tête (L8).")
            st.stop()

        df_full_read["Stock"] = pd.to_numeric(df_full_read["Stock"], errors='coerce').fillna(0)
        df_full_read["Tarif d'achat"] = pd.to_numeric(df_full_read["Tarif d'achat"], errors='coerce').fillna(0)
        df_full_read["Conditionnement"] = pd.to_numeric(df_full_read["Conditionnement"], errors='coerce').fillna(1).apply(lambda x: int(x) if x > 0 else 1)
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

        df_full_state = st.session_state.df_full
        # --- Corrected Filter Logic ---
        df_init_filt_temp_read = df_full_state[
            (df_full_state["Fournisseur"].astype(str).str.strip() != "") &
            (df_full_state["Fournisseur"].astype(str).str.strip().str.lower() != "#filter") & # Corrected .str.lower()
            (df_full_state["AF_RefFourniss"].astype(str).str.strip() != "")
        ].copy()
        st.session_state.df_initial_filtered = df_init_filt_temp_read

        first_week_col_idx_approx = 12
        potential_sem_cols_read = []
        if len(df_full_state.columns) > first_week_col_idx_approx:
            candidate_cols_sem = df_full_state.columns[first_week_col_idx_approx:].tolist()
            known_non_week_cols_set = set(["Tarif d'achat", "Conditionnement", "Stock", "Total", "Stock à terme", "Ventes N-1", "Ventes 12 semaines identiques N-1", "Ventes 12 dernières semaines", "Quantité à commander", "Fournisseur", "AF_RefFourniss", "Référence Article", "Désignation Article"])
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

        if not df_init_filt_temp_read.empty: st.session_state.unique_suppliers_list = sorted(df_init_filt_temp_read["Fournisseur"].astype(str).unique().tolist())
        else: st.session_state.unique_suppliers_list = []
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
                       "Vérification Stock", "Simulation Forecast", "Suivi Commandes Fourn."]
    tab1, tab1_ai, tab2, tab3, tab4, tab5 = st.tabs(tab_titles_main)

    # --- Tab 1: Classic Order Forecast ---
    with tab1:
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
                                        lbl_c_s_idx_exp=get_column_letter(exp_c_s_c_exp.index(lbl_name_col_exp)+1)

                                        total_row_xl_idx_exp=n_r_exp+2
                                        ws_exp[f"{lbl_c_s_idx_exp}{total_row_xl_idx_exp}"]="TOTAL";ws_exp[f"{lbl_c_s_idx_exp}{total_row_xl_idx_exp}"].font=Font(bold=True)
                                        cell_gt_exp=ws_exp[f"{t_l_exp}{total_row_xl_idx_exp}"]
                                        if n_r_exp>0:cell_gt_exp.value=f"=SUM({t_l_exp}2:{t_l_exp}{n_r_exp+1})"
                                        else:cell_gt_exp.value=0
                                        cell_gt_exp.number_format='#,##0.00€';cell_gt_exp.font=Font(bold=True)

                                        min_req_row_xl_idx_exp=n_r_exp+3
                                        ws_exp[f"{lbl_c_s_idx_exp}{min_req_row_xl_idx_exp}"]="Min Requis Fourn.";ws_exp[f"{lbl_c_s_idx_exp}{min_req_row_xl_idx_exp}"].font=Font(bold=True)
                                        cell_min_req_v_exp=ws_exp[f"{t_l_exp}{min_req_row_xl_idx_exp}"]
                                        min_r_s_val_exp=min_o_amts.get(sup_e_exp,0);min_d_s_val_exp=f"{min_r_s_val_exp:,.2f}€"if min_r_s_val_exp>0 else"N/A"
                                        cell_min_req_v_exp.value=min_d_s_val_exp;cell_min_req_v_exp.font=Font(bold=True)
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

            # --- Display Current Stock Value ---
            if sel_f_t1_ai and not df_disp_t1_ai.empty:
                try:
                    # Use double quotes for column names with special characters or spaces
                    stock_actuel_selection_ai = pd.to_numeric(df_disp_t1_ai["Stock"], errors='coerce').fillna(0)
                    tarif_achat_selection_ai = pd.to_numeric(df_disp_t1_ai["Tarif d'achat"], errors='coerce').fillna(0)
                    valeur_stock_selection_ai = (stock_actuel_selection_ai * tarif_achat_selection_ai).sum()
                    st.metric(label="📊 Valeur Stock Actuel (€) (Fourn. Sél.)", value=f"{valeur_stock_selection_ai:,.2f} €")
                except KeyError as e_stockval:
                    st.error(f"Erreur : Colonne manquante pour valeur stock ('{e_stockval}').")
                except Exception as e_stockval_calc:
                    st.error(f"Erreur calcul valeur stock actuel : {e_stockval_calc}")
            elif sel_f_t1_ai and df_disp_t1_ai.empty:
                 st.metric(label="📊 Valeur Stock Actuel (€) (Fourn. Sél.)", value="0,00 €")
            # --- End Display Current Stock Value ---
            else:
                st.info("Sélectionner fournisseur(s).")

            st.markdown("---")

            if df_disp_t1_ai.empty and sel_f_t1_ai:
                st.warning("Aucun article trouvé pour le(s) fournisseur(s) sélectionné(s).")
            elif not id_sem_cols and not df_disp_t1_ai.empty:
                st.warning("Colonnes ventes historiques non identifiées. Prévision IA impossible.")
            elif not df_disp_t1_ai.empty:
                st.markdown("#### Paramètres Prévision IA")
                c1_ai, c2_ai, c3_ai = st.columns(3)
                with c1_ai:
                    fcst_w_ai_t1 = st.number_input("⏳ Semaines à prévoir:", 1, 52, value=st.session_state.ai_forecast_weeks_val, step=1, key="fcst_w_ai_t1_numin")
                with c2_ai:
                    min_amt_ai_t1_default = st.session_state.ai_min_order_val
                    if len(sel_f_t1_ai) == 1 and sel_f_t1_ai[0] in min_o_amts and min_amt_ai_t1_default == 0.0:
                        min_amt_ai_t1_default = min_o_amts[sel_f_t1_ai[0]]
                    min_amt_ai_t1 = st.number_input("💶 Montant min (€) (si 1 fourn.):", 0.0, value=min_amt_ai_t1_default, step=50.0, format="%.2f", key="min_amt_ai_t1_numin")
                with c3_ai:
                    default_reduc_target = st.session_state.get('ai_stock_reduc_target_val', 0.0)
                    stock_reduc_target_ai_t1 = st.number_input(
                        "📉 Objectif Réduc. Val. Stock (€):",
                        min_value=0.0,
                        value=default_reduc_target,
                        step=100.0,
                        format="%.2f",
                        key="stock_reduc_target_ai_t1_numin",
                        help="Entrez le montant dont vous souhaitez réduire la valeur globale du stock projeté (pour les articles commandés). La commande sera réduite pour tenter d'atteindre cet objectif, au risque d'augmenter les ruptures."
                    )

                st.session_state.ai_forecast_weeks_val = fcst_w_ai_t1
                st.session_state.ai_min_order_val = min_amt_ai_t1
                st.session_state.ai_stock_reduc_target_val = stock_reduc_target_ai_t1

                if st.button("🚀 Calculer Qtés avec IA", key="calc_q_ai_b_t1_go"):
                    curr_calc_params_t1_ai = {
                        'suppliers': sel_f_t1_ai,
                        'forecast_weeks': fcst_w_ai_t1,
                        'min_amount_ui': min_amt_ai_t1,
                        'stock_reduc_target': stock_reduc_target_ai_t1,
                        'sem_cols_hash': hash(tuple(id_sem_cols))
                    }
                    st.session_state.ai_commande_params_calculated_for = curr_calc_params_t1_ai

                    res_dfs_list_ai_calc = []
                    calc_ok_overall_ai = True
                    st.info(f"Lancement prévision IA pour {len(sel_f_t1_ai)} fournisseur(s)...")

                    # --- Calculation Loop ---
                    for sup_idx_ai, sup_name_proc_ai in enumerate(sel_f_t1_ai):
                        # Progress update can be shown here if needed
                        df_sup_subset_ai_proc = df_disp_t1_ai[df_disp_t1_ai["Fournisseur"] == sup_name_proc_ai].copy()
                        sup_specific_min_order_ai = min_amt_ai_t1 if len(sel_f_t1_ai) == 1 else min_o_amts.get(sup_name_proc_ai, 0.0)

                        if not df_sup_subset_ai_proc.empty:
                            ai_res_df_sup, _ = ai_calculate_order_quantities( # Ignore amount returned here, recalc later
                                df_sup_subset_ai_proc,
                                id_sem_cols,
                                fcst_w_ai_t1,
                                sup_specific_min_order_ai
                            )
                            if ai_res_df_sup is not None:
                                res_dfs_list_ai_calc.append(ai_res_df_sup)
                            else:
                                st.error(f"Échec calcul IA pour: {sup_name_proc_ai}")
                                calc_ok_overall_ai = False
                        else: logging.info(f"Aucun article pour {sup_name_proc_ai} (IA).")
                    # --- End Calculation Loop ---

                    # --- Process Results ---
                    df_after_reduction_filter = pd.DataFrame()

                    if calc_ok_overall_ai and res_dfs_list_ai_calc:
                        final_ai_res_df_calc = pd.concat(res_dfs_list_ai_calc, ignore_index=True) if res_dfs_list_ai_calc else pd.DataFrame()
                        st.success("✅ Calcul IA initial terminé!")

                        # --- Apply 350€ Filter ---
                        st.markdown("---")
                        st.info("Application du filtre : Commandes fournisseur < 350€ ignorées (sauf si article en stock < 0).")
                        df_after_350_filter = pd.DataFrame()
                        if not final_ai_res_df_calc.empty:
                            # Ensure Total Cmd is numeric before grouping
                            final_ai_res_df_calc['Total Cmd (€) (IA)'] = pd.to_numeric(final_ai_res_df_calc['Total Cmd (€) (IA)'], errors='coerce').fillna(0)
                            final_ai_res_df_calc['Qté Cmdée (IA)'] = pd.to_numeric(final_ai_res_df_calc['Qté Cmdée (IA)'], errors='coerce').fillna(0)
                            final_ai_res_df_calc['Stock'] = pd.to_numeric(final_ai_res_df_calc['Stock'], errors='coerce').fillna(0)

                            order_value_per_supplier = final_ai_res_df_calc[final_ai_res_df_calc['Qté Cmdée (IA)'] > 0].groupby('Fournisseur')['Total Cmd (€) (IA)'].sum()
                            suppliers_with_neg_stock_ordered = final_ai_res_df_calc[(final_ai_res_df_calc['Qté Cmdée (IA)'] > 0) & (final_ai_res_df_calc['Stock'] < 0)]['Fournisseur'].unique()
                            suppliers_to_keep = set(s for s, v in order_value_per_supplier.items() if v >= 350 or s in suppliers_with_neg_stock_ordered)

                            initial_rows_350 = len(final_ai_res_df_calc)
                            df_after_350_filter = final_ai_res_df_calc[final_ai_res_df_calc['Fournisseur'].isin(suppliers_to_keep)].copy()
                            filtered_rows_350 = len(df_after_350_filter)
                            if initial_rows_350 > filtered_rows_350: st.caption(f"{initial_rows_350 - filtered_rows_350} lignes article (< 350€ sans stock négatif) retirées.")
                        else:
                             df_after_350_filter = final_ai_res_df_calc

                        # --- Apply Stock Reduction Filter (Low Rotation Strategy) ---
                        df_after_reduction_filter = df_after_350_filter.copy()
                        reduction_target_value = st.session_state.ai_stock_reduc_target_val

                        if reduction_target_value > 0 and not df_after_reduction_filter.empty:
                            st.markdown("---")
                            st.info(f"Tentative de réduction (-{reduction_target_value:,.2f}€) en ciblant faible rotation...")

                            # Calculate WoS for prioritization
                            try:
                                wos_period_weeks = 12
                                available_weeks = len(id_sem_cols)
                                weeks_to_use_for_wos = min(wos_period_weeks, available_weeks)
                                if weeks_to_use_for_wos > 0:
                                    semaine_cols_for_wos = id_sem_cols[-weeks_to_use_for_wos:]
                                    valid_wos_cols = [col for col in semaine_cols_for_wos if col in df_after_reduction_filter.columns]
                                    if valid_wos_cols:
                                        for col in valid_wos_cols: df_after_reduction_filter[col] = pd.to_numeric(df_after_reduction_filter[col], errors='coerce').fillna(0)
                                        avg_weekly_sales = df_after_reduction_filter[valid_wos_cols].sum(axis=1) / weeks_to_use_for_wos
                                        current_stock_wos = pd.to_numeric(df_after_reduction_filter['Stock'], errors='coerce').fillna(0)
                                        df_after_reduction_filter['WoS_Calculated'] = np.divide(current_stock_wos, avg_weekly_sales, out=np.full_like(current_stock_wos, np.inf, dtype=np.float64), where=avg_weekly_sales != 0)
                                        df_after_reduction_filter.loc[current_stock_wos <= 0, 'WoS_Calculated'] = 0.0
                                        st.caption(f"Priorisation réduction basée sur WoS ({weeks_to_use_for_wos} sem.)")
                                    else:
                                        df_after_reduction_filter['WoS_Calculated'] = np.inf; st.warning("Impossible de calculer WoS (cols ventes récentes manquantes).")
                                else:
                                    df_after_reduction_filter['WoS_Calculated'] = np.inf; st.warning("Pas assez d'historique pour WoS.")
                            except Exception as e_wos:
                                st.error(f"Erreur calcul WoS: {e_wos}"); df_after_reduction_filter['WoS_Calculated'] = np.inf

                            # Ensure other columns are numeric
                            # --- Corrected Line (Using double quotes) ---
                            for col in ["Tarif d'achat", "Qté Cmdée (IA)", "Conditionnement", "Stock"]:
                                if col in df_after_reduction_filter.columns:
                                    df_after_reduction_filter[col] = pd.to_numeric(df_after_reduction_filter[col], errors='coerce').fillna(0)
                            # --- End Correction ---
                            df_after_reduction_filter['Conditionnement'] = df_after_reduction_filter['Conditionnement'].apply(lambda x: int(x) if x > 0 else 1)
                            df_after_reduction_filter['Qté Cmdée (IA)'] = df_after_reduction_filter['Qté Cmdée (IA)'].astype(int)

                            # Reduction Logic
                            current_stock_value_reduc = (df_after_reduction_filter['Stock'] * df_after_reduction_filter['Tarif d'achat']).sum()
                            target_max_stock_value_reduc = max(0, current_stock_value_reduc - reduction_target_value)
                            projected_stock_value_reduc = ((df_after_reduction_filter['Stock'] + df_after_reduction_filter['Qté Cmdée (IA)']) * df_after_reduction_filter['Tarif d'achat']).sum()
                            value_to_reduce_reduc = max(0, projected_stock_value_reduc - target_max_stock_value_reduc)

                            st.caption(f"Val. Stock Actuel (Cmd Filt.): {current_stock_value_reduc:,.2f}€ | Val. Stock Projeté (Cmd Filt.): {projected_stock_value_reduc:,.2f}€ | Val. Cible Max: {target_max_stock_value_reduc:,.2f}€")

                            if value_to_reduce_reduc > 0.01:
                                logging.info(f"Objectif réduction stock: Excédent de {value_to_reduce_reduc:,.2f}€ à réduire.")
                                candidates_reduc = df_after_reduction_filter[df_after_reduction_filter['Qté Cmdée (IA)'] > 0].copy()

                                if not candidates_reduc.empty:
                                    candidates_reduc.sort_values(by='WoS_Calculated', ascending=False, inplace=True, na_position='first')
                                    value_reduced_total = 0.0
                                    max_loops_reduc = len(candidates_reduc) * 10
                                    loops_count_reduc = 0
                                    candidate_indices_reduc = candidates_reduc.index.tolist()

                                    while value_to_reduce_reduc > 0.01 and loops_count_reduc < max_loops_reduc:
                                        made_reduction_this_pass = False
                                        for item_index_reduc in candidate_indices_reduc:
                                            loops_count_reduc += 1
                                            if value_to_reduce_reduc <= 0.01 or loops_count_reduc >= max_loops_reduc: break

                                            current_qty_reduc = df_after_reduction_filter.loc[item_index_reduc, 'Qté Cmdée (IA)']
                                            packaging_reduc = df_after_reduction_filter.loc[item_index_reduc, 'Conditionnement']
                                            price_reduc = df_after_reduction_filter.loc[item_index_reduc, 'Tarif d'achat']

                                            if packaging_reduc > 0 and price_reduc > 0 and current_qty_reduc >= packaging_reduc:
                                                value_per_pkg_reduc = packaging_reduc * price_reduc
                                                if value_to_reduce_reduc >= value_per_pkg_reduc * 0.5:
                                                    df_after_reduction_filter.loc[item_index_reduc, 'Qté Cmdée (IA)'] -= packaging_reduc
                                                    value_to_reduce_reduc -= value_per_pkg_reduc
                                                    value_reduced_total += value_per_pkg_reduc
                                                    made_reduction_this_pass = True
                                                    logging.debug(f"Réduit Qty index {item_index_reduc} (WoS: {df_after_reduction_filter.loc[item_index_reduc, 'WoS_Calculated']:.1f}) by {packaging_reduc}. Reste: {value_to_reduce_reduc:.2f}€")
                                                    if value_to_reduce_reduc <= 0.01: break

                                        if not made_reduction_this_pass or loops_count_reduc >= max_loops_reduc :
                                            break # Exit outer loop if no progress or max loops

                                    st.caption(f"Réduction appliquée: {value_reduced_total:,.2f}€ retirés de la commande.")
                                    if value_to_reduce_reduc > 0.01: st.warning(f"Objectif réduction non atteint. Reste {value_to_reduce_reduc:,.2f}€ excédent.")

                                    # Recalculate final derived columns
                                    df_after_reduction_filter['Total Cmd (€) (IA)'] = df_after_reduction_filter['Qté Cmdée (IA)'] * df_after_reduction_filter['Tarif d'achat']
                                    df_after_reduction_filter['Stock Terme (IA)'] = df_after_reduction_filter['Stock'] + df_after_reduction_filter['Qté Cmdée (IA)']
                                else:
                                    st.caption("Aucun article commandé trouvé pour appliquer la réduction.")
                            else:
                                st.caption("Aucune réduction nécessaire pour objectif stock.")
                        else:
                            st.caption("Pas d'objectif de réduction de stock spécifié ou pas de données après filtre 350€.")
                        # --- End Stock Reduction Filter ---

                        # Final state update
                        st.session_state.ai_commande_result_df = df_after_reduction_filter
                        st.session_state.ai_commande_total_amount = df_after_reduction_filter['Total Cmd (€) (IA)'].sum() if not df_after_reduction_filter.empty else 0.0

                        st.rerun()

                    elif not res_dfs_list_ai_calc:
                        st.error("❌ Aucun résultat IA n'a pu être généré.")
                        st.session_state.ai_commande_result_df = pd.DataFrame()
                        st.session_state.ai_commande_total_amount = 0.0
                    else: # Partial success
                        st.warning("Certains calculs IA ont échoué. Filtre 350€ appliqué, filtre réduction stock non appliqué sur résultats partiels.")
                        df_after_350_filter = pd.DataFrame()
                        if res_dfs_list_ai_calc:
                           final_ai_res_df_calc = pd.concat(res_dfs_list_ai_calc, ignore_index=True) if res_dfs_list_ai_calc else pd.DataFrame()
                           if not final_ai_res_df_calc.empty:
                               # Ensure numeric before processing
                               final_ai_res_df_calc['Total Cmd (€) (IA)'] = pd.to_numeric(final_ai_res_df_calc['Total Cmd (€) (IA)'], errors='coerce').fillna(0)
                               final_ai_res_df_calc['Qté Cmdée (IA)'] = pd.to_numeric(final_ai_res_df_calc['Qté Cmdée (IA)'], errors='coerce').fillna(0)
                               final_ai_res_df_calc['Stock'] = pd.to_numeric(final_ai_res_df_calc['Stock'], errors='coerce').fillna(0)

                               order_value_per_supplier = final_ai_res_df_calc[final_ai_res_df_calc['Qté Cmdée (IA)'] > 0].groupby('Fournisseur')['Total Cmd (€) (IA)'].sum()
                               suppliers_with_neg_stock_ordered = final_ai_res_df_calc[(final_ai_res_df_calc['Qté Cmdée (IA)'] > 0) & (final_ai_res_df_calc['Stock'] < 0)]['Fournisseur'].unique()
                               suppliers_to_keep = set(s for s, v in order_value_per_supplier.items() if v >= 350 or s in suppliers_with_neg_stock_ordered)
                               df_after_350_filter = final_ai_res_df_calc[final_ai_res_df_calc['Fournisseur'].isin(suppliers_to_keep)].copy()
                           else:
                               df_after_350_filter = final_ai_res_df_calc

                        st.session_state.ai_commande_result_df = df_after_350_filter
                        st.session_state.ai_commande_total_amount = df_after_350_filter['Total Cmd (€) (IA)'].sum() if not df_after_350_filter.empty else 0.0
                        st.rerun()

                # --- Display Results ---
                if 'ai_commande_result_df' in st.session_state and st.session_state.ai_commande_result_df is not None:
                    curr_ui_params_t1_ai_disp = {
                        'suppliers': sel_f_t1_ai,
                        'forecast_weeks': fcst_w_ai_t1,
                        'min_amount_ui': min_amt_ai_t1,
                        'stock_reduc_target': stock_reduc_target_ai_t1,
                        'sem_cols_hash': hash(tuple(id_sem_cols))
                    }
                    if st.session_state.get('ai_commande_params_calculated_for') == curr_ui_params_t1_ai_disp:
                        st.markdown("---")
                        st.markdown("#### Résultats Prévision Commande (IA) - *Ajustés si nécessaire*")
                        df_disp_ai_res_final = st.session_state.ai_commande_result_df
                        total_amt_ai_res_final = st.session_state.ai_commande_total_amount

                        st.metric(label="💰 Montant Total Cmd Final (€) (IA)", value=f"{total_amt_ai_res_final:,.2f} €")

                        if not df_disp_ai_res_final.empty:
                            # Ensure numeric for final display calculation
                            df_disp_ai_res_final['Stock'] = pd.to_numeric(df_disp_ai_res_final['Stock'], errors='coerce').fillna(0)
                            df_disp_ai_res_final['Qté Cmdée (IA)'] = pd.to_numeric(df_disp_ai_res_final['Qté Cmdée (IA)'], errors='coerce').fillna(0)
                            df_disp_ai_res_final['Tarif d'achat'] = pd.to_numeric(df_disp_ai_res_final['Tarif d'achat'], errors='coerce').fillna(0)
                            final_proj_stock_value = ((df_disp_ai_res_final['Stock'] + df_disp_ai_res_final['Qté Cmdée (IA)']) * df_disp_ai_res_final['Tarif d'achat']).sum()
                            st.metric(label="📊 Valeur Stock Projeté Final (€) (Articles Commandés)", value=f"{final_proj_stock_value:,.2f} €")

                        for sup_chk_min_ai in sel_f_t1_ai:
                            sup_min_cfg_val_ai = min_o_amts.get(sup_chk_min_ai, 0.0)
                            min_applied_in_calc_ai = min_amt_ai_t1 if len(sel_f_t1_ai) == 1 else sup_min_cfg_val_ai
                            if min_applied_in_calc_ai > 0 and not df_disp_ai_res_final.empty: # Check if DF not empty
                                actual_order_sup_ai = df_disp_ai_res_final[(df_disp_ai_res_final["Fournisseur"] == sup_chk_min_ai)]["Total Cmd (€) (IA)"].sum()
                                if actual_order_sup_ai < min_applied_in_calc_ai:
                                    st.warning(f"⚠️ Min cmd pour {sup_chk_min_ai} ({min_applied_in_calc_ai:,.2f}€) non atteint ({actual_order_sup_ai:,.2f}€) - *peut être dû à la réduction de stock*.")

                        cols_show_ai_res_final = ["Fournisseur","AF_RefFourniss","Référence Article","Désignation Article", "Stock", "Forecast Ventes (IA)"]
                        if 'WoS_Calculated' in df_disp_ai_res_final.columns: cols_show_ai_res_final.append("WoS_Calculated")
                        cols_show_ai_res_final.extend(["Conditionnement", "Qté Cmdée (IA)", "Stock Terme (IA)", "Tarif d'achat", "Total Cmd (€) (IA)"])
                        disp_cols_ai_final = [c for c in cols_show_ai_res_final if c in df_disp_ai_res_final.columns]

                        if not disp_cols_ai_final: st.error("Aucune col à afficher (résultats IA).")
                        else:
                            fmts_ai_final = {"Tarif d'achat":"{:,.2f}€","Total Cmd (€) (IA)":"{:,.2f}€","Forecast Ventes (IA)":"{:,.2f}","Stock":"{:,.0f}","Conditionnement":"{:,.0f}","Qté Cmdée (IA)":"{:,.0f}","Stock Terme (IA)":"{:,.0f}"}
                            if 'WoS_Calculated' in disp_cols_ai_final: fmts_ai_final["WoS_Calculated"] = "{:,.1f}"

                            df_display_ordered_only = df_disp_ai_res_final[df_disp_ai_res_final["Qté Cmdée (IA)"] > 0] if "Qté Cmdée (IA)" in df_disp_ai_res_final else df_disp_ai_res_final

                            if df_display_ordered_only.empty and not df_disp_ai_res_final.empty:
                                st.info("Aucune quantité à commander après application des filtres et objectifs.")
                            elif not df_display_ordered_only.empty :
                                df_display_styled = df_display_ordered_only[disp_cols_ai_final].copy()
                                if 'WoS_Calculated' in df_display_styled: df_display_styled['WoS_Calculated'] = df_display_styled['WoS_Calculated'].replace([np.inf, -np.inf], ">999")
                                st.dataframe(df_display_styled.style.format(fmts_ai_final,na_rep="-",thousands=","))
                            else:
                                st.dataframe(df_disp_ai_res_final[disp_cols_ai_final].style.format(fmts_ai_final,na_rep="-",thousands=","))

                        # Export Final Adjusted Results
                        st.markdown("#### Export Commandes Finales (IA)")
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
                                        if 'WoS_Calculated' in df_w_s_ai_dl.columns:
                                            df_w_s_ai_dl['WoS_Calculated'] = df_w_s_ai_dl['WoS_Calculated'].replace([np.inf, -np.inf], 9999)

                                        n_r_ai_dl=len(df_w_s_ai_dl);s_nm_ai_dl=sanitize_sheet_name(f"IA_Cmd_{sup_e_ai_dl}")
                                        df_w_s_ai_dl.to_excel(writer_ai_exp_dl,sheet_name=s_nm_ai_dl,index=False)
                                        ws_ai_dl=writer_ai_exp_dl.sheets[s_nm_ai_dl]
                                        cmd_col_fmts_ai_dl={"Stock":"#,##0","Forecast Ventes (IA)":"#,##0.00","Conditionnement":"#,##0","Qté Cmdée (IA)":"#,##0","Stock Terme (IA)":"#,##0","Tarif d'achat":"#,##0.00€"}
                                        if 'WoS_Calculated' in exp_cols_sheet_ai_dl:
                                            cmd_col_fmts_ai_dl["WoS_Calculated"] = "0.0"

                                        format_excel_sheet(ws_ai_dl,df_w_s_ai_dl,column_formats=cmd_col_fmts_ai_dl)
                                        if f_ok_ai_dl and n_r_ai_dl>0:
                                            for r_idx_ai_dl in range(2,n_r_ai_dl+2):cell_t_ai_dl=ws_ai_dl[f"{t_l_ai_dl}{r_idx_ai_dl}"];cell_t_ai_dl.value=f"={q_l_ai_dl}{r_idx_ai_dl}*{p_l_ai_dl}{r_idx_ai_dl}";cell_t_ai_dl.number_format='#,##0.00€'

                                        lbl_name_col_ai_dl="Désignation Article"
                                        if lbl_name_col_ai_dl not in exp_cols_sheet_ai_dl: lbl_name_col_ai_dl = exp_cols_sheet_ai_dl[1] if len(exp_cols_sheet_ai_dl)>1 else exp_cols_sheet_ai_dl[0]
                                        lbl_c_s_idx_ai_dl=get_column_letter(exp_cols_sheet_ai_dl.index(lbl_name_col_ai_dl)+1)

                                        total_row_xl_idx_ai_dl=n_r_ai_dl+2
                                        ws_ai_dl[f"{lbl_c_s_idx_ai_dl}{total_row_xl_idx_ai_dl}"]="TOTAL";ws_ai_dl[f"{lbl_c_s_idx_ai_dl}{total_row_xl_idx_ai_dl}"].font=Font(bold=True)
                                        cell_gt_ai_dl=ws_ai_dl[f"{t_l_ai_dl}{total_row_xl_idx_ai_dl}"]
                                        if n_r_ai_dl>0:cell_gt_ai_dl.value=f"=SUM({t_l_ai_dl}2:{t_l_ai_dl}{n_r_ai_dl+1})"
                                        else:cell_gt_ai_dl.value=0
                                        cell_gt_ai_dl.number_format='#,##0.00€';cell_gt_ai_dl.font=Font(bold=True)

                                        min_req_row_xl_idx_ai_dl=n_r_ai_dl+3
                                        ws_ai_dl[f"{lbl_c_s_idx_ai_dl}{min_req_row_xl_idx_ai_dl}"]="Min Requis Fourn.";ws_ai_dl[f"{lbl_c_s_idx_ai_dl}{min_req_row_xl_idx_ai_dl}"].font=Font(bold=True)
                                        cell_min_req_v_ai_dl=ws_ai_dl[f"{t_l_ai_dl}{min_req_row_xl_idx_ai_dl}"]
                                        min_r_s_val_ai_dl=min_o_amts.get(sup_e_ai_dl,0);min_d_s_val_ai_dl=f"{min_r_s_val_ai_dl:,.2f}€"if min_r_s_val_ai_dl>0 else"N/A"
                                        cell_min_req_v_ai_dl.value=min_d_s_val_ai_dl;cell_min_req_v_ai_dl.font=Font(bold=True)
                                        shts_ai_exp_dl+=1
                                if shts_ai_exp_dl > 0:
                                    out_b_ai_exp_dl.seek(0)
                                    fn_ai_dl=f"commandes_IA_ajustees_{'multi'if len(sel_f_t1_ai)>1 else sanitize_sheet_name(sel_f_t1_ai[0])}_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx"
                                    st.download_button(f"📥 Télécharger Commandes Finales ({shts_ai_exp_dl} feuilles)",out_b_ai_exp_dl,fn_ai_dl,"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",key="dl_ai_cmd_final_b_t1_dl")
                                else:st.info("Aucune qté IA > 0 à exporter après application des filtres et objectifs.")
                            except Exception as e_wrt_ai_dl:logging.exception(f"Err ExcelWriter cmd IA ajustée: {e_wrt_ai_dl}");st.error("Erreur export commandes IA finales.")
                        else:st.info("Aucun article qté IA > 0 à exporter après application des filtres et objectifs.")
                    else:st.info("Paramètres IA ou objectif de réduction changés. Relancer calcul pour résultats à jour.")

    # --- Tab 2: Stock Rotation Analysis ---
    with tab2:
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
                if df_r_res_t2 is not None:
                    st.success("✅ Analyse rotation OK.");st.session_state.rotation_result_df=df_r_res_t2
                    st.rerun()
                else:st.error("❌ Analyse rotation échouée.")

            if st.session_state.rotation_result_df is not None:
                curr_ui_params_t2_disp = {'suppliers': sel_f_t2, 'period_label': sel_p_lbl_r_t2, 'show_all': show_all_r_t2, 'threshold': r_thr_ui_t2 if not show_all_r_t2 else -1, 'sem_cols_hash': hash(tuple(id_sem_cols))}
                if st.session_state.get('rotation_params_calculated_for') == curr_ui_params_t2_disp:
                    st.markdown("---");st.markdown(f"#### Résultats Rotation ({sel_p_lbl_r_t2})")
                    df_r_orig_t2=st.session_state.rotation_result_df

                    df_r_disp_t2_final=pd.DataFrame();df_r_to_fmt_t2_final=pd.DataFrame()
                    if df_r_orig_t2.empty:st.info("Aucune donnée rotation à afficher.")
                    elif show_all_r_t2:
                        df_r_disp_t2_final=df_r_orig_t2.copy();df_r_to_fmt_t2_final=df_r_disp_t2_final.copy();st.caption(f"Affichage {len(df_r_disp_t2_final)} articles.")
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
                    if df_f_res_t4 is not None:
                        st.success("✅ Simu forecast OK.");st.session_state.forecast_result_df=df_f_res_t4;st.session_state.forecast_grand_total_amount=gt_f_t4
                        st.rerun()
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

# Fallback if no file is uploaded or if data loading failed and state was reset
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
