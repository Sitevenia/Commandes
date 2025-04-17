
import pandas as pd
from datetime import datetime
import numpy as np

def extract_week_number(week_str):
    try:
        return int(week_str.split('-S')[1])
    except:
        return None

def run_forecast_simulation(df):
    week_cols = [col for col in df.columns if '-S' in col]
    current_week_num = datetime.today().isocalendar().week
    valid_weeks = [col for col in week_cols if extract_week_number(col) is not None and extract_week_number(col) <= current_week_num]
    
    if not valid_weeks:
        raise ValueError("Aucune semaine valide détectée dans les colonnes.")

    current_week = max(valid_weeks, key=lambda x: extract_week_number(x))

    results = []
    for _, row in df.iterrows():
        stock = row["Stock"]
        conditionnement = row["Conditionnement"]
        quantite_mini = row.get("Quantité mini", 0)
        
        sales_data = row[valid_weeks[-12:]].values if len(valid_weeks) >= 12 else row[valid_weeks].values
        avg_sales = sum(sales_data) / len(sales_data) if len(sales_data) > 0 else 0
        
        if quantite_mini == 0:
            qty_needed = max(0, -stock)
        else:
            qty_needed = max(0, quantite_mini - stock)
        
        if qty_needed > 0:
            qty_final = int(((qty_needed - 1) // conditionnement + 1) * conditionnement)
        else:
            qty_final = 0

        results.append(qty_final)
    
    df["Quantité commandée"] = results
    return df


def run_target_stock_sim(df, valeur_stock_cible):
    df = df.copy()
    df["Valeur stock actuel"] = df["Stock"] * df["Tarif d’achat"]
    valeur_actuelle = df["Valeur stock actuel"].sum()

    df["Commande proposée"] = 0

    # Calcul d'un besoin brut par produit
    for i, row in df.iterrows():
        stock = row["Stock"]
        mini = row.get("Quantité mini", 0)
        cond = row["Conditionnement"]
        tarif = row["Tarif d’achat"]

        if mini == 0 and stock < 0:
            besoin = -stock
        elif mini > 0 and stock < mini:
            besoin = mini - stock
        else:
            besoin = 0

        if besoin > 0:
            qte = int(np.ceil(besoin / cond)) * cond
        else:
            qte = 0

        df.at[i, "Commande proposée"] = qte

    df["Valeur ajoutée"] = df["Commande proposée"] * df["Tarif d’achat"]
    df = df.sort_values(by="Valeur ajoutée", ascending=False)

    total = valeur_actuelle
    df["Quantité commandée"] = 0

    for i, row in df.iterrows():
        if total >= valeur_stock_cible:
            break
        qte = row["Commande proposée"]
        cond = row["Conditionnement"]
        prix = row["Tarif d’achat"]

        while qte >= cond:
            if total + (cond * prix) <= valeur_stock_cible:
                df.at[i, "Quantité commandée"] += cond
                total += cond * prix
                qte -= cond
            else:
                break

    return df
