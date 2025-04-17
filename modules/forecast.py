
import pandas as pd
from datetime import datetime

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
    # Calcul de la valeur de stock actuelle par produit
    df["Valeur stock"] = df["Stock"] * df["Tarif d’achat"]
    valeur_stock_actuelle = df["Valeur stock"].sum()
    
    if valeur_stock_actuelle <= valeur_stock_cible:
        df["Quantité commandée"] = 0
        return df

    facteur_reduction = valeur_stock_cible / valeur_stock_actuelle
    
    results = []
    for _, row in df.iterrows():
        stock = row["Stock"]
        conditionnement = row["Conditionnement"]
        tarif = row["Tarif d’achat"]
        quantite_mini = row.get("Quantité mini", 0)
        
        # Stock visé par produit
        valeur_stock_produit = stock * tarif
        valeur_cible_produit = valeur_stock_produit * facteur_reduction
        stock_cible = valeur_cible_produit / tarif
        qty_needed = max(0, quantite_mini - stock_cible) if quantite_mini > 0 else max(0, -stock)

        # Respect du conditionnement
        if qty_needed > 0:
            qty_final = int(((qty_needed - 1) // conditionnement + 1) * conditionnement)
        else:
            qty_final = 0
        
        results.append(qty_final)
    
    df["Quantité commandée"] = results
    return df
