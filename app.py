import streamlit as st
import pandas as pd
import numpy as np
import io
import logging # Optional: for better debugging if needed

# Setup basic logging (optional)
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def safe_read_excel(uploaded_file, sheet_name, **kwargs):
    """Safely reads an Excel sheet, returning None if sheet not found."""
    try:
        return pd.read_excel(uploaded_file, sheet_name=sheet_name, **kwargs)
    except ValueError as e:
        # ValueError can be raised if sheet_name doesn't exist
        logging.warning(f"Sheet '{sheet_name}' not found or error reading it: {e}")
        st.warning(f"⚠️ L'onglet '{sheet_name}' n'a pas été trouvé dans le fichier Excel. Les vérifications associées seront ignorées.")
        return None
    except Exception as e:
        logging.error(f"Unexpected error reading sheet '{sheet_name}': {e}")
        st.error(f"❌ Erreur inattendue lors de la lecture de l'onglet '{sheet_name}'.")
        return None

def calculer_quantite_a_commander(df, semaine_columns, montant_minimum_input, duree_semaines):
    """
    Calcule la quantité à commander pour chaque produit en fonction des ventes passées,
    du stock actuel, du conditionnement et d'un montant minimum de commande (fourni en entrée).

    Args:
        df (pd.DataFrame): DataFrame contenant les données des produits.
        semaine_columns (list): Liste des noms des colonnes représentant les ventes hebdomadaires.
        montant_minimum_input (float): Montant minimum de commande fourni par l'utilisateur via l'interface.
                                      La fonction utilisera cette valeur pour tenter d'ajuster les quantités.
        duree_semaines (int): Nombre de semaines de ventes à couvrir par la commande.

    Returns:
        tuple: (quantite_a_commander, ventes_N1, ventes_12_N1, ventes_12_last, montant_total_final)
               ou None en cas d'erreur.
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
        df_calc = df.copy() # Work on a copy to avoid modifying the original df passed to the function
        for col in required_cols:
            df_calc[col] = pd.to_numeric(df_calc[col], errors='coerce').replace([np.inf, -np.inf], np.nan).fillna(0)

        # --- Calculs des Ventes Moyennes ---
        num_semaines_totales = len(semaine_columns)
        ventes_N1 = df_calc[semaine_columns].sum(axis=1)

        # Calculs N-1 (avec garde-fous pour le nombre de semaines)
        if num_semaines_totales >= 64:
            ventes_12_semaines_N1 = df_calc[semaine_columns[-64:-52]].sum(axis=1)
            ventes_12_semaines_N1_suivantes = df_calc[semaine_columns[-52:-40]].sum(axis=1)
            avg_12_N1 = ventes_12_semaines_N1 / 12
            avg_12_N1_suivantes = ventes_12_semaines_N1_suivantes / 12
        else:
            # st.warning("Données N-1 (< 64 semaines) insuffisantes, pondération ajustée.")
            ventes_12_semaines_N1 = pd.Series(0, index=df_calc.index)
            ventes_12_semaines_N1_suivantes = pd.Series(0, index=df_calc.index)
            avg_12_N1 = 0
            avg_12_N1_suivantes = 0

        # Calculs 12 dernières semaines (avec garde-fous)
        nb_semaines_recentes = min(num_semaines_totales, 12)
        if nb_semaines_recentes > 0:
            ventes_12_dernieres_semaines = df_calc[semaine_columns[-nb_semaines_recentes:]].sum(axis=1)
            avg_12_dernieres = ventes_12_dernieres_semaines / nb_semaines_recentes
        else:
            # st.warning("Aucune donnée de vente récente (< 12 semaines), pondération ajustée.")
            ventes_12_dernieres_semaines = pd.Series(0, index=df_calc.index)
            avg_12_dernieres = 0


        # --- Calcul de la Quantité Pondérée ---
        quantite_ponderee = (0.5 * avg_12_dernieres +
                             0.2 * avg_12_N1 +
                             0.3 * avg_12_N1_suivantes)

        quantite_necessaire = quantite_ponderee * duree_semaines
        quantite_a_commander_series = (quantite_necessaire - df_calc["Stock"]).apply(lambda x: max(0, x))

        # --- Ajustements Basés sur les Règles ---
        conditionnement = df_calc["Conditionnement"]
        stock_actuel = df_calc["Stock"]
        tarif_achat = df_calc["Tarif d'achat"]
        quantite_a_commander = quantite_a_commander_series.tolist()

        # Ajuster aux conditionnements
        for i in range(len(quantite_a_commander)):
            cond = conditionnement.iloc[i]
            q = quantite_a_commander[i]
            if q > 0 and cond > 0:
                quantite_a_commander[i] = int(np.ceil(q / cond) * cond)
            elif q > 0 and cond <= 0:
                 # Log warning if needed, set to 0
                 quantite_a_commander[i] = 0
            else:
                quantite_a_commander[i] = 0

        # Règle 1: Vendu >= 2 fois (12 dernières) ET Stock <= 1 => Min 1 cond.
        if nb_semaines_recentes > 0:
            for i in range(len(quantite_a_commander)):
                cond = conditionnement.iloc[i]
                ventes_recentes_count = (df_calc[semaine_columns[-nb_semaines_recentes:]].iloc[i] > 0).sum()
                if ventes_recentes_count >= 2 and stock_actuel.iloc[i] <= 1 and cond > 0:
                    quantite_a_commander[i] = max(quantite_a_commander[i], cond)

        # Règle 2: Ventes N-1 < 6 ET Ventes 12 dernières < 2 => Qte = 0
        for i in range(len(quantite_a_commander)):
            ventes_tot_n1 = ventes_N1.iloc[i]
            ventes_recentes_sum = ventes_12_dernieres_semaines.iloc[i]
            if ventes_tot_n1 < 6 and ventes_recentes_sum < 2:
                quantite_a_commander[i] = 0

        # --- Ajustement pour Montant Minimum (basé sur l'INPUT utilisateur) ---
        montant_total_avant_ajust_min = sum(q * p for q, p in zip(quantite_a_commander, tarif_achat))

        # Utiliser la valeur du champ 'montant_minimum_input' fournie à la fonction
        if montant_minimum_input > 0 and montant_total_avant_ajust_min < montant_minimum_input:
            montant_total_actuel = montant_total_avant_ajust_min
            indices_commandes = [i for i, q in enumerate(quantite_a_commander) if q > 0]
            # Optional: Sort indices by price * conditionnement descending to reach minimum faster?
            # indices_commandes.sort(key=lambda i: tarif_achat.iloc[i] * conditionnement.iloc[i] if conditionnement.iloc[i]>0 else 0, reverse=True)

            idx_pointer = 0
            max_iterations = len(df_calc) * 10 # Safety break for infinite loops
            iterations = 0

            while montant_total_actuel < montant_minimum_input and iterations < max_iterations:
                iterations += 1
                if not indices_commandes:
                    logging.warning(f"Impossible d'atteindre le montant minimum input de {montant_minimum_input:.2f}€ car aucune quantité initiale commandée. Montant actuel: {montant_total_actuel:.2f}€")
                    break

                current_idx = indices_commandes[idx_pointer % len(indices_commandes)]
                cond = conditionnement.iloc[current_idx]
                prix = tarif_achat.iloc[current_idx]

                if cond > 0 and prix > 0:
                    quantite_a_commander[current_idx] += cond
                    montant_total_actuel += cond * prix
                elif cond <= 0 :
                    # Remove index if conditionnement is invalid to avoid infinite loop on this item
                    logging.warning(f"Conditionnement invalide (<=0) pour produit index {current_idx} lors de l'ajustement du min. Ignoré.")
                    indices_commandes.pop(idx_pointer % len(indices_commandes))
                    if not indices_commandes: continue
                    idx_pointer -= 1 # Adjust pointer as list size changed

                idx_pointer += 1

            if iterations >= max_iterations and montant_total_actuel < montant_minimum_input:
                 logging.error(f"Ajustement du montant minimum ({montant_minimum_input:.2f}€) échoué après {max_iterations} itérations. Montant atteint: {montant_total_actuel:.2f}€. Vérifiez conditionnements/prix.")
                 st.error("L'ajustement automatique pour atteindre le montant minimum a échoué (possible boucle ou conditionnements/prix nuls).")


        # Recalculer le montant final après tous les ajustements
        montant_total_final = sum(q * p for q, p in zip(quantite_a_commander, tarif_achat))

        return (quantite_a_commander,
                ventes_N1,
                ventes_12_semaines_N1,
                ventes_12_dernieres_semaines,
                montant_total_final) # Retourner le montant final calculé

    except KeyError as e:
        st.error(f"Erreur de clé: Colonne '{e}' introuvable pendant le calcul.")
        logging.error(f"KeyError during calculation: {e}")
        return None
    except ValueError as e:
         st.error(f"Erreur de valeur: Problème avec les données numériques pendant le calcul - {e}")
         logging.error(f"ValueError during calculation: {e}")
         return None
    except Exception as e:
        st.error(f"Erreur inattendue dans le calcul des quantités: {e}")
        logging.exception("Unexpected error during quantity calculation:") # Log full traceback
        return None


# --- Configuration de la page Streamlit ---
st.set_page_config(page_title="Forecast App", layout="wide")
st.title("📦 Application de Prévision des Commandes")

# --- Chargement du fichier ---
uploaded_file = st.file_uploader("📁 Charger le fichier Excel principal", type=["xlsx", "xls"])

# Initialize variables to store dataframes and minimums
df_full = None
df_min_commande = None
min_order_dict = {}

if uploaded_file:
    # Use BytesIO to read the file into memory once, avoiding re-reading
    file_buffer = io.BytesIO(uploaded_file.getvalue())

    # --- Read Main Sheet ("Tableau final") ---
    logging.info("Attempting to read 'Tableau final' sheet.")
    df_full = safe_read_excel(file_buffer, sheet_name="Tableau final", header=7)

    # --- Read Minimum Order Sheet ("Minimum de commande") ---
    logging.info("Attempting to read 'Minimum de commande' sheet.")
    # Make sure the buffer position is reset if reading multiple sheets from buffer
    file_buffer.seek(0)
    df_min_commande = safe_read_excel(file_buffer, sheet_name="Minimum de commande") # Assume header is row 1 (index 0) by default

    if df_min_commande is not None:
        # --- Process Minimum Order Data ---
        logging.info("Processing 'Minimum de commande' sheet.")
        # **Important**: Adjust column names based on your actual Excel file
        supplier_col_min = "Fournisseur" # Column name for supplier in 'Minimum de commande' sheet
        min_amount_col = "Minimum de Commande" # Column name for minimum amount

        required_min_cols = [supplier_col_min, min_amount_col]
        if all(col in df_min_commande.columns for col in required_min_cols):
            try:
                # Clean supplier names and convert minimums to numeric
                df_min_commande[supplier_col_min] = df_min_commande[supplier_col_min].astype(str).str.strip()
                df_min_commande[min_amount_col] = pd.to_numeric(df_min_commande[min_amount_col], errors='coerce')

                # Create the dictionary, dropping rows where minimum is NaN or supplier is empty
                min_order_dict = df_min_commande.dropna(subset=[supplier_col_min, min_amount_col])\
                                               .set_index(supplier_col_min)[min_amount_col]\
                                               .to_dict()
                logging.info(f"Successfully created minimum order dictionary with {len(min_order_dict)} entries.")
                # st.write("Minimums lus:", min_order_dict) # Debug line
            except KeyError as e:
                st.error(f"❌ Colonne attendue '{e}' non trouvée dans l'onglet 'Minimum de commande'. Vérifiez les noms des colonnes.")
                logging.error(f"KeyError processing minimum order sheet: {e}")
                min_order_dict = {} # Reset dict if error
            except Exception as e:
                 st.error(f"❌ Erreur lors du traitement de l'onglet 'Minimum de commande': {e}")
                 logging.exception("Error processing minimum order sheet:")
                 min_order_dict = {} # Reset dict if error
        else:
            missing_min_cols = [col for col in required_min_cols if col not in df_min_commande.columns]
            st.warning(f"⚠️ Colonnes requises ({', '.join(missing_min_cols)}) manquantes dans l'onglet 'Minimum de commande'. La vérification des minimums ne peut pas être effectuée.")
            logging.warning(f"Missing required columns in 'Minimum de commande' sheet: {missing_min_cols}")


    # --- Continue with Main Processing only if 'Tableau final' was read successfully ---
    if df_full is not None:
        st.success("✅ Fichier principal ('Tableau final') chargé.")

        # --- Initial Filtering (Suppliers, Refs) ---
        try:
            df = df_full[
                (df_full["Fournisseur"].notna()) &
                (df_full["Fournisseur"] != "") &
                (df_full["Fournisseur"] != "#FILTER") &
                (df_full["AF_RefFourniss"].notna()) &
                (df_full["AF_RefFourniss"] != "")
            ].copy()

            if df.empty:
                 st.warning("Aucune ligne valide trouvée après le filtrage initial (Fournisseur/AF_RefFourniss renseignés et non '#FILTER').")
                 fournisseurs = []
            else:
                fournisseurs = sorted(df["Fournisseur"].unique().tolist())

        except KeyError as e:
            st.error(f"❌ Colonne essentielle '{e}' manquante dans 'Tableau final' pour le filtrage initial.")
            st.stop() # Stop execution if basic filtering fails


        # --- User Selection ---
        selected_fournisseurs = st.multiselect(
            "👤 Sélectionnez le(s) fournisseur(s)",
            options=fournisseurs,
            default=[]
        )

        # Filter data based on selection
        if selected_fournisseurs:
            df_filtered = df[df["Fournisseur"].isin(selected_fournisseurs)].copy()
        else:
            df_filtered = pd.DataFrame(columns=df.columns)


        # --- Identify Week Columns and Prepare Data ---
        start_col_index = 12 # Index de la colonne "N"
        semaine_columns = []
        if len(df_filtered.columns) > start_col_index:
            potential_week_cols = df_filtered.columns[start_col_index:].tolist()
            exclude_columns = ["Tarif d'achat", "Conditionnement", "Stock", "Total", "Stock à terme",
                               "Ventes N-1", "Ventes 12 semaines identiques N-1", "Ventes 12 dernières semaines",
                               "Quantité à commander"]

            semaine_columns = [
                col for col in potential_week_cols
                if col not in exclude_columns and pd.api.types.is_numeric_dtype(df_filtered.get(col, pd.Series(dtype=float)).dtype)
            ] # Added .get() for safety if column disappears after filtering

            if not semaine_columns:
                 st.warning("⚠️ Aucune colonne numérique de ventes hebdomadaires identifiée après l'index 12.")

            # Ensure essential numeric columns exist and are numeric
            essential_numeric_cols = ["Stock", "Conditionnement", "Tarif d'achat"]
            for col in essential_numeric_cols:
                 if col in df_filtered.columns:
                     df_filtered[col] = pd.to_numeric(df_filtered[col], errors='coerce').fillna(0)
                 elif not df_filtered.empty: # Only error if df is supposed to have data
                     st.error(f"Colonne essentielle '{col}' manquante dans les données filtrées.")
                     st.stop() # Stop if essential calculation columns are missing

        elif not df_filtered.empty:
            st.warning("Le fichier ne semble pas contenir de colonnes après l'index 12 pour les données de ventes.")


        # --- Calculation Parameters ---
        col1, col2 = st.columns(2)
        with col1:
            duree_semaines = st.number_input("⏳ Durée de couverture (semaines)", value=4, min_value=1, step=1, key="duree_input")
        with col2:
            # The key allows us to potentially update this value programmatically later if needed, though the request is for user update
            montant_minimum_input_val = st.number_input(
                "💶 Montant minimum de commande (€)",
                value=0.0, min_value=0.0, step=50.0, format="%.2f",
                key="montant_min_input",
                help="Montant utilisé pour l'ajustement des quantités. Si un seul fournisseur est sélectionné, une alerte peut suggérer de modifier cette valeur si elle est inférieure au minimum requis dans l'onglet 'Minimum de commande'."
            )

        # --- Execute Calculation and Display ---
        if not df_filtered.empty and semaine_columns:
            st.info("🚀 Lancement du calcul...")
            result = calculer_quantite_a_commander(
                df_filtered,
                semaine_columns,
                montant_minimum_input_val, # Pass the user input value here
                duree_semaines
            )

            if result is not None:
                st.success("✅ Calculs terminés.")
                (quantite_calculée, ventes_N1_calc, ventes_12_N1_calc,
                 ventes_12_last_calc, montant_total_calc) = result

                # Add results to DataFrame
                df_filtered.loc[:, "Quantité à commander"] = quantite_calculée
                df_filtered.loc[:, "Ventes N-1"] = ventes_N1_calc
                df_filtered.loc[:, "Ventes 12 semaines identiques N-1"] = ventes_12_N1_calc
                df_filtered.loc[:, "Ventes 12 dernières semaines"] = ventes_12_last_calc
                df_filtered.loc[:, "Total"] = df_filtered["Tarif d'achat"] * df_filtered["Quantité à commander"]
                df_filtered.loc[:, "Stock à terme"] = df_filtered["Stock"] + df_filtered["Quantité à commander"]

                # --- Display Metrics and Potential Minimum Order Warning ---
                st.metric(label="💰 Montant total de la commande calculée", value=f"{montant_total_calc:.2f} €")

                # **NEW**: Check against minimum from Excel sheet if *one* supplier selected
                if len(selected_fournisseurs) == 1:
                    selected_supplier = selected_fournisseurs[0]
                    if selected_supplier in min_order_dict:
                        required_minimum = min_order_dict[selected_supplier]
                        if required_minimum > 0 and montant_total_calc < required_minimum:
                            diff = required_minimum - montant_total_calc
                            st.warning(
                                f"⚠️ **Minimum de Commande Non Atteint!**\n"
                                f"Fournisseur: **{selected_supplier}**\n"
                                f"Montant Calculé: **{montant_total_calc:.2f} €**\n"
                                f"Minimum Requis (fichier Excel): **{required_minimum:.2f} €** (Manque: {diff:.2f} €)\n\n"
                                f"➡️ **Suggestion:** Pour tenter d'atteindre ce minimum, vous pouvez modifier le champ "
                                f"'Montant minimum de commande (€)' ci-dessus et le définir à **{required_minimum:.2f}**, "
                                f"puis relancer le calcul (l'application essaiera d'ajouter des produits pour atteindre ce seuil)."
                            )
                        # Optional: Add an info message if the minimum was met or exceeded ?
                        # elif required_minimum > 0 and montant_total_calc >= required_minimum:
                        #    st.success(f"✅ Le minimum de commande ({required_minimum:.2f}€) pour {selected_supplier} est atteint ou dépassé.")

                elif len(selected_fournisseurs) > 1 and any(f in min_order_dict and min_order_dict[f] > 0 for f in selected_fournisseurs):
                    # Inform user that individual minimums aren't checked for multi-selection
                     st.info("ℹ️ Vérification individuelle des minimums de commande désactivée lors de la sélection de plusieurs fournisseurs.")


                # --- Display Results Table ---
                st.subheader("📊 Résultats de la prévision")
                required_display_columns = ["AF_RefFourniss", "Référence Article", "Désignation Article", "Stock"]
                display_columns_base = required_display_columns + [
                    "Ventes N-1", "Ventes 12 semaines identiques N-1", "Ventes 12 dernières semaines",
                    "Conditionnement", "Quantité à commander", "Stock à terme",
                    "Tarif d'achat", "Total"
                ]
                display_columns = [col for col in display_columns_base if col in df_filtered.columns]
                missing_display_columns = [col for col in required_display_columns if col not in df_filtered.columns]

                if missing_display_columns:
                    st.error(f"❌ Colonnes manquantes pour l'affichage : {', '.join(missing_display_columns)}")
                else:
                    st.dataframe(df_filtered[display_columns].style.format({
                        "Tarif d'achat": "{:.2f}€", "Total": "{:.2f}€",
                        "Ventes N-1": "{:,.0f}", "Ventes 12 semaines identiques N-1": "{:,.0f}",
                        "Ventes 12 dernières semaines": "{:,.0f}", "Stock": "{:,.0f}",
                        "Conditionnement": "{:,.0f}", "Quantité à commander": "{:,.0f}",
                        "Stock à terme": "{:,.0f}"
                    }, na_rep="-")) # Added thousands separator and na_rep

                # --- Export ---
                st.subheader("⬇️ Exportation")
                df_export = df_filtered[df_filtered["Quantité à commander"] > 0].copy()

                if not df_export.empty:
                    export_columns = display_columns # Use same columns as displayed
                    df_export_final = df_export[export_columns].copy()

                    # Add total row
                    total_row_dict = {col: "" for col in export_columns}
                    total_row_dict["Désignation Article"] = "TOTAL COMMANDE"
                    total_row_dict["Total"] = df_export_final["Total"].sum()
                    total_row_df = pd.DataFrame([total_row_dict])
                    df_with_total = pd.concat([df_export_final, total_row_df], ignore_index=True)

                    # Prepare Excel file in memory
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine="openpyxl") as writer:
                        df_with_total.to_excel(writer, sheet_name="Quantités_à_commander", index=False)
                    output.seek(0)

                    suppliers_str = "_".join(selected_fournisseurs).replace(" ", "_").replace("/", "-")[:50] # Limit filename length
                    filename = f"commande_{suppliers_str}_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx"

                    st.download_button(
                        label="📥 Télécharger le fichier de commande (Excel)",
                        data=output,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.info("ℹ️ Aucune quantité à commander pour l'exportation avec les paramètres actuels.")

            else:
                # Error handled and displayed within calculer_quantite_a_commander
                st.error("❌ Le calcul n'a pas pu aboutir. Vérifiez les messages d'erreur.")

        # --- Conditions for no calculation ---
        elif not selected_fournisseurs:
            st.warning("⚠️ Veuillez sélectionner au moins un fournisseur.")
        elif not semaine_columns and not df.empty: # df is not empty, but no week columns found
             st.warning("⚠️ Calcul impossible: aucune colonne de ventes hebdomadaires n'a été identifiée ou les données filtrées sont incomplètes.")
        # Add elif df_filtered.empty and selected_fournisseurs:?

    # --- Error Handling for File Loading ---
    elif uploaded_file and df_full is None: # File was uploaded but reading 'Tableau final' failed
        st.error("❌ Échec de la lecture de l'onglet 'Tableau final'. Vérifiez le nom de l'onglet et le format du fichier.")
    elif not uploaded_file:
        st.info("👋 Bienvenue ! Chargez votre fichier Excel pour commencer.")

# General error catch at the end (less likely needed with specific catches)
# except Exception as e:
#    st.error(f"❌ Une erreur générale est survenue dans l'application : {e}")
#    logging.exception("Unhandled exception in Streamlit app:")
