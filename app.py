# --- (Previous code: imports, functions, data loading, calculation, display) ---

                # --- EXPORT LOGIC (Revised for Robust Formulas) ---
                st.subheader("⬇️ Exportation Excel par Fournisseur (avec formules)")
                df_export_all = df_filtered[df_filtered["Quantité à commander"] > 0].copy()

                if not df_export_all.empty:
                    output = io.BytesIO()
                    sheets_created_count = 0
                    try:
                        # Use ExcelWriter context manager
                        with pd.ExcelWriter(output, engine="openpyxl") as writer:
                            logging.info(f"Export: Starting for suppliers: {selected_fournisseurs}")

                            # --- Define column names (CRITICAL: Must match DataFrame columns) ---
                            qty_col_name = "Quantité à commander"
                            price_col_name = "Tarif d'achat"
                            total_col_name = "Total"
                            # Define columns for export sheets - includes columns needed for formula AND display
                            export_columns = [col for col in display_columns if col != 'Fournisseur' and col in df_export_all.columns]

                            # --- Verify essential columns exist and get letters ---
                            formula_ready = False
                            if not all(c in export_columns for c in [qty_col_name, price_col_name, total_col_name]):
                                st.error(f"❌ Export Error: Columns '{qty_col_name}', '{price_col_name}', or '{total_col_name}' not found in export columns.")
                                logging.error(f"Export Error: Essential formula columns missing from {export_columns}")
                            else:
                                try:
                                    # Get 0-based indices relative to the export_columns list
                                    qty_col_idx = export_columns.index(qty_col_name)
                                    price_col_idx = export_columns.index(price_col_name)
                                    total_col_idx = export_columns.index(total_col_name)
                                    # Convert to 1-based Excel column letters
                                    qty_col_letter = get_column_letter(qty_col_idx + 1)
                                    price_col_letter = get_column_letter(price_col_idx + 1)
                                    total_col_letter = get_column_letter(total_col_idx + 1)
                                    formula_ready = True
                                    logging.info(f"Export: Formula cols identified: Qty={qty_col_letter}, Price={price_col_letter}, Total={total_col_letter}")
                                except ValueError as e:
                                    st.error(f"❌ Export Error: Could not find index for formula columns: {e}")
                                    logging.error(f"Export Error: Could not get column index: {e}")
                                except Exception as e_idx:
                                     st.error(f"❌ Export Error: Unexpected error getting column indices: {e_idx}")
                                     logging.exception("Export Error: Getting column indices failed.")


                            # --- Proceed only if columns for formula are identified ---
                            if formula_ready:
                                for supplier in selected_fournisseurs:
                                    logging.info(f"Export: Processing supplier sheet for {supplier}")
                                    df_supplier_export = df_export_all[df_export_all["Fournisseur"] == supplier].copy() # Use copy

                                    if not df_supplier_export.empty:
                                        # Keep the calculated 'Total' column for now, it will be overwritten by formulas
                                        df_supplier_sheet_data = df_supplier_export[export_columns].copy()
                                        num_data_rows = len(df_supplier_sheet_data)
                                        logging.info(f"Export: {supplier} - Found {num_data_rows} data rows.")

                                        # --- Summary Rows Prep ---
                                        # Note: supplier_total_val is now just for reference, the sheet will use SUM
                                        supplier_total_val = df_supplier_sheet_data[total_col_name].sum()
                                        required_minimum = min_order_dict.get(supplier, 0)
                                        min_formatted = f"{required_minimum:,.2f} €" if required_minimum > 0 else "N/A"
                                        label_col = "Désignation Article" if "Désignation Article" in export_columns else export_columns[1]
                                        value_col_for_summary = total_col_name # Column where SUM and Min text go

                                        # Create DataFrames for summary rows (Total value is placeholder)
                                        total_row_dict = {col: "" for col in export_columns}; total_row_dict[label_col] = "TOTAL COMMANDE"; total_row_dict[value_col_for_summary] = supplier_total_val # Placeholder
                                        min_row_dict = {col: "" for col in export_columns}; min_row_dict[label_col] = "Minimum Requis"; min_row_dict[value_col_for_summary] = min_formatted
                                        total_row_df = pd.DataFrame([total_row_dict]); min_row_df = pd.DataFrame([min_row_dict])

                                        # Concatenate data + summary rows
                                        df_sheet = pd.concat([df_supplier_sheet_data, total_row_df, min_row_df], ignore_index=True)

                                        sanitized_name = sanitize_sheet_name(supplier)
                                        try:
                                            # --- Step 1: Write DataFrame to Excel (includes calculated values) ---
                                            logging.debug(f"Export: Writing df_sheet to sheet '{sanitized_name}'")
                                            df_sheet.to_excel(writer, sheet_name=sanitized_name, index=False)

                                            # --- Step 2: Get the openpyxl worksheet object ---
                                            worksheet = writer.sheets[sanitized_name]
                                            logging.debug(f"Export: Got worksheet object for '{sanitized_name}'")

                                            # --- Step 3: Overwrite 'Total' cells with FORMULAS ---
                                            logging.debug(f"Export: Applying formulas to rows 2 to {num_data_rows + 1} in col {total_col_letter}")
                                            # Excel rows are 1-based. Header is 1. Data starts row 2.
                                            for excel_row_num in range(2, num_data_rows + 2): # +2 because range end is exclusive
                                                # Construct the formula string for this row
                                                formula_str = f"={qty_col_letter}{excel_row_num}*{price_col_letter}{excel_row_num}"
                                                # Get the cell object
                                                cell = worksheet[f"{total_col_letter}{excel_row_num}"]
                                                # Assign the formula string to the cell's value
                                                cell.value = formula_str
                                                # Apply number format (optional, but good)
                                                cell.number_format = '#,##0.00 €'
                                                # logging.debug(f"Export: Wrote formula '{formula_str}' to {cell.coordinate}") # Verbose logging

                                            # --- Step 4: Apply SUM formula to the grand total row ---
                                            total_formula_row_num = num_data_rows + 2 # Excel row number for "TOTAL COMMANDE"
                                            logging.debug(f"Export: Applying SUM formula to row {total_formula_row_num} in col {total_col_letter}")
                                            if num_data_rows > 0: # Only add sum if there was data
                                                # Construct SUM formula
                                                sum_formula_str = f"=SUM({total_col_letter}2:{total_col_letter}{num_data_rows + 1})"
                                                # Get the cell for the SUM
                                                sum_cell = worksheet[f"{total_col_letter}{total_formula_row_num}"]
                                                # Assign the formula string
                                                sum_cell.value = sum_formula_str
                                                # Apply formatting
                                                sum_cell.number_format = '#,##0.00 €'
                                                # Optional: Apply bold font to summary rows here if needed (using openpyxl Font)
                                                # from openpyxl.styles import Font
                                                # bold_font = Font(bold=True)
                                                # label_col_letter = get_column_letter(export_columns.index(label_col) + 1)
                                                # min_req_row_num = total_formula_row_num + 1
                                                # worksheet[f"{label_col_letter}{total_formula_row_num}"].font = bold_font
                                                # worksheet[f"{total_col_letter}{total_formula_row_num}"].font = bold_font
                                                # worksheet[f"{label_col_letter}{min_req_row_num}"].font = bold_font
                                                # worksheet[f"{total_col_letter}{min_req_row_num}"].font = bold_font

                                            sheets_created_count += 1
                                            logging.info(f"Export: Successfully processed sheet '{sanitized_name}' with formulas.")

                                        except Exception as write_error:
                                            st.error(f"❌ Erreur écriture/formules pour {supplier} ({sanitized_name}): {write_error}")
                                            logging.exception(f"Export: Error writing sheet/formulas for {sanitized_name}:") # Log full traceback

                                    else: # df_supplier_export was empty
                                         logging.info(f"Export: No items to order for supplier '{supplier}', skipping sheet creation.")
                                # End of loop for suppliers
                            else: # Formula not ready
                                 st.error("❌ Export annulé car les colonnes nécessaires pour les formules n'ont pas pu être identifiées.")
                    except Exception as e_writer:
                        st.error(f"❌ Erreur majeure lors de la création du fichier Excel : {e_writer}")
                        logging.exception("Export: Error during ExcelWriter context:")

                    # --- Download Button ---
                    if sheets_created_count > 0:
                        output.seek(0)
                        suppliers_str = "multiples" if len(selected_fournisseurs) > 1 else sanitize_sheet_name(selected_fournisseurs[0])
                        timestamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M')
                        filename = f"commande_{suppliers_str}_{timestamp}.xlsx"
                        st.download_button(label=f"📥 Télécharger Commandes ({sheets_created_count} Onglet{'s' if sheets_created_count > 1 else ''})", data=output, file_name=filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                        logging.info(f"Export: Download button created for {sheets_created_count} sheets.")
                    elif formula_ready: # Formulas were intended but no data rows existed
                         st.info("ℹ️ Aucune quantité à commander trouvée pour l'exportation.")
                         logging.info("Export: No sheets created (no qty > 0).")

                else: # df_export_all was empty
                    st.info("ℹ️ Aucune quantité à commander globale trouvée pour l'exportation.")
                    logging.info("Export: df_export_all was empty.")

            else: # Calculation result was None
                st.error("❌ Le calcul n'a pas pu aboutir.")

        # --- Conditions for no calculation ---
        elif not selected_fournisseurs: st.warning("⚠️ Veuillez sélectionner au moins un fournisseur.")
        elif not semaine_columns and not df_filtered.empty: st.warning("⚠️ Calcul impossible: pas de colonnes ventes valides.")

# --- App footer/initial message ---
elif not uploaded_file:
    st.info("👋 Bienvenue ! Chargez votre fichier Excel.")
