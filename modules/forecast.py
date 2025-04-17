
import pandas as pd
import numpy as np
from datetime import datetime

def extract_week_number(week_str):
    try:
        return int(week_str.split('-S')[1])
    except:
        return None

def run_forecast_simulation(df):
    df = df.copy()
    df["Valeur stock actuel"] = df["Stock"] * df["Tarif d’achat"]

    week_cols = [col for col in df.columns if '-S' in col]
    current_week_num = datetime.today().isocalendar().week
    valid_weeks = [col for col in week_cols if extract_week_number(col) is not None and extract_week_number(col) <= current_week_num]

    if not valid_weeks:
        raise ValueError("Aucune semaine valide détectée dans les colonnes.")

    quantites = []
    valeurs_ajoutees = []
    valeurs_totales = []

    for _, row in df.iterrows():
        stock = row["Stock"]
        cond = row["Conditionnement"]
        mini = row.get("Quantité mini", 0)
        prix = row["Tarif d’achat"]
        valeur_stock = stock * prix

        sales_data = row[valid_weeks[-12:]].values if len(valid_weeks) >= 12 else row[valid_weeks].values
        avg_sales = sum(sales_data) / len(sales_data) if len(sales_data) > 0 else 0

        if mini == 0:
            qty_needed = max(0, -stock)
        else:
            qty_needed = max(0, mini - stock)

        if qty_needed > 0:
            qte = int(np.ceil(qty_needed / cond)) * cond
        else:
            qte = 0

        valeur_ajout = qte * prix
        valeur_totale = valeur_stock + valeur_ajout

        quantites.append(qte)
        valeurs_ajoutees.append(valeur_ajout)
        valeurs_totales.append(valeur_totale)

    df["Quantité commandée"] = quantites
    df["Valeur ajoutée"] = valeurs_ajoutees
    df["Valeur totale"] = valeurs_totales

    return df

def run_target_stock_sim(df, valeur_stock_cible):
    df = df.copy()
    df["Valeur stock actuel"] = df["Stock"] * df["Tarif d’achat"]

    week_cols = [col for col in df.columns if '-S' in col]
    current_week = datetime.today().isocalendar().week
    valid_weeks = [col for col in week_cols if extract_week_number(col) is not None and extract_week_number(col) <= current_week]
    ventes_12s = df[valid_weeks[-12:]].mean(axis=1) if len(valid_weeks) >= 12 else df[valid_weeks].mean(axis=1)

    df["Ventes moyennes 12s"] = ventes_12s
    df["Quantité commandée"] = 0
    df["Valeur ajoutée"] = 0.0
    df["Valeur totale"] = df["Valeur stock actuel"]

    df_valid = df[df["Ventes moyennes 12s"] > 0].copy()

    if df_valid.empty:
        df["Stock total après commande"] = df["Stock"]
        return df

    score_total = df_valid["Ventes moyennes 12s"].sum()
    df_valid["Poids"] = df_valid["Ventes moyennes 12s"] / score_total

    valeur_stock_actuelle = df["Valeur stock actuel"].sum()
    montant_a_ajouter = valeur_stock_cible - valeur_stock_actuelle

    for i, row in df_valid.iterrows():
        poids = row["Poids"]
        part_montant = montant_a_ajouter * poids
        prix = row["Tarif d’achat"]
        cond = row["Conditionnement"]

        qte = int(np.floor(part_montant / prix / cond)) * cond
        valeur_ajout = qte * prix
        valeur_totale = row["Valeur stock actuel"] + valeur_ajout

        df_valid.at[i, "Quantité commandée"] = qte
        df_valid.at[i, "Valeur ajoutée"] = valeur_ajout
        df_valid.at[i, "Valeur totale"] = valeur_totale

    df.update(df_valid[["Quantité commandée", "Valeur ajoutée", "Valeur totale"]])
    df["Stock total après commande"] = df["Stock"] + df["Quantité commandée"]

    total_row = pd.DataFrame({
        "Produit": ["TOTAL"],
        "Stock": [df["Stock"].sum()],
        "Conditionnement": [""],
        "Tarif d’achat": [""],
        "Quantité mini": [""],
        "Valeur stock actuel": [df["Valeur stock actuel"].sum()],
        "Quantité commandée": [df["Quantité commandée"].sum()],
        "Valeur ajoutée": [df["Valeur ajoutée"].sum()],
        "Valeur totale": [df["Valeur totale"].sum()],
        "Stock total après commande": [df["Stock total après commande"].sum()]
    })

    df.drop(columns=["Ventes moyennes 12s"], inplace=True)
    df = pd.concat([df, total_row], ignore_index=True)

    return df
