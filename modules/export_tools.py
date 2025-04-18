
def export_three_sheets(df_standard, df_target):
    import pandas as pd
    import io
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill

    output = io.BytesIO()

    # Création du comparatif
    df_comp = df_standard.copy()
    df_comp = df_comp.rename(columns={col: f"{col} (standard)" for col in df_comp.columns if col not in ['Produit', 'Désignation']})
    df_target_renamed = df_target.copy()
    df_target_renamed = df_target_renamed.rename(columns={col: f"{col} (objectif)" for col in df_target.columns if col not in ['Produit', 'Désignation']})

    df_comparatif = pd.merge(df_comp, df_target_renamed, on=['Produit', 'Désignation'], how='outer')

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_standard.to_excel(writer, sheet_name="Simulation standard", index=False)
        df_target.to_excel(writer, sheet_name="Simulation objectif", index=False)
        df_comparatif.to_excel(writer, sheet_name="Comparatif", index=False)
    return output
