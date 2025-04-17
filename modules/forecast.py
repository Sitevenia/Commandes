
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

    df["Quantité commandée"] = 0
    df["Valeur ajoutée"] = 0.0
    df["Valeur totale"] = df["Valeur stock actuel"]

    df["Score produit"] = ventes_12s
    df = df.sort_values(by="Score produit", ascending=False).reset_index(drop=True)

    def calculer_valeur_totale(df_):
        return df_["Valeur totale"].sum()

    max_iterations = 10000
    iterations = 0

    while calculer_valeur_totale(df) < valeur_stock_cible and iterations < max_iterations:
        for i, row in df.iterrows():
            cond = row["Conditionnement"]
            prix = row["Tarif d’achat"]
            mini = row.get("Quantité mini", 0)
            stock = row["Stock"]
            qte_actuelle = df.at[i, "Quantité commandée"]

            # Conditions pour être éligible à l'ajustement
            if mini == 0 and stock >= 0 and qte_actuelle > 0:
                continue
            if mini > 0 and (stock + qte_actuelle) >= mini:
                continue

            # Ajouter une unité de conditionnement
            df.at[i, "Quantité commandée"] += cond
            df.at[i, "Valeur ajoutée"] = df.at[i, "Quantité commandée"] * prix
            df.at[i, "Valeur totale"] = df.at[i, "Valeur stock actuel"] + df.at[i, "Valeur ajoutée"]

            if calculer_valeur_totale(df) >= valeur_stock_cible:
                break

        iterations += 1

    df.drop(columns=["Score produit"], inplace=True)
    return df
