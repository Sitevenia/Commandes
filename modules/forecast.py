
import pandas as pd

def run_forecast_simulation(df):
    df = df.copy()

    if "Valeur stock actuel" not in df.columns:
        df["Valeur stock actuel"] = df["Stock"] * df["Tarif d’achat"]

    df["Quantité commandée"] = 0
    df["Valeur ajoutée"] = 0.0
    df["Valeur totale"] = df["Valeur stock actuel"]
    df["Stock total après commande"] = df["Stock"]

    conditionnement = df["Conditionnement"].replace(0, 1)
    df["Tarif d’achat"].fillna(0, inplace=True)

    # Simulation simple : logique métier basée sur ventes faibles ou stock bas
    for idx in df.index:
        stock = df.at[idx, "Stock"]
        quantite_mini = df.at[idx, "Quantité mini"] if "Quantité mini" in df.columns else 0

        if quantite_mini == 0:
            if stock < 0:
                qte = abs(stock)
                qte += conditionnement[idx] - (qte % conditionnement[idx]) if qte % conditionnement[idx] != 0 else 0
                df.at[idx, "Quantité commandée"] = qte
        else:
            if stock < quantite_mini:
                qte = quantite_mini - stock
                qte += conditionnement[idx] - (qte % conditionnement[idx]) if qte % conditionnement[idx] != 0 else 0
                df.at[idx, "Quantité commandée"] = qte

        df.at[idx, "Valeur ajoutée"] = df.at[idx, "Quantité commandée"] * df.at[idx, "Tarif d’achat"]
        df.at[idx, "Valeur totale"] = df.at[idx, "Valeur stock actuel"] + df.at[idx, "Valeur ajoutée"]
        df.at[idx, "Stock total après commande"] = df.at[idx, "Stock"] + df.at[idx, "Quantité commandée"]

    if "Produit" in df.columns:
        total_row = {col: df[col].sum() if pd.api.types.is_numeric_dtype(df[col]) else "" for col in df.columns}
        total_row["Produit"] = "TOTAL"
        df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)

    return df

def run_target_stock_sim(df, objectif):
    df = df.copy()

    if "Valeur stock actuel" not in df.columns:
        df["Valeur stock actuel"] = df["Stock"] * df["Tarif d’achat"]

    df["Quantité commandée"] = 0
    df["Valeur ajoutée"] = 0.0
    df["Valeur totale"] = df["Valeur stock actuel"]
    df["Stock total après commande"] = df["Stock"]

    conditionnement = df["Conditionnement"].replace(0, 1)
    df["Tarif d’achat"].fillna(0, inplace=True)

    produits = df[df["Tarif d’achat"] > 0].copy()
    produits["Score"] = (
        produits["Valeur stock actuel"] + 1
    ) / (produits["Tarif d’achat"] + 1)

    produits = produits.sort_values(by="Score")
    i = 0

    while df["Valeur totale"].sum() < objectif and not produits.empty:
        idx = produits.index[i]
        df.at[idx, "Quantité commandée"] += conditionnement[idx]
        df.at[idx, "Valeur ajoutée"] = df.at[idx, "Quantité commandée"] * df.at[idx, "Tarif d’achat"]
        df.at[idx, "Valeur totale"] = df.at[idx, "Valeur stock actuel"] + df.at[idx, "Valeur ajoutée"]
        df.at[idx, "Stock total après commande"] = df.at[idx, "Stock"] + df.at[idx, "Quantité commandée"]

        i = (i + 1) % len(produits)
        if df["Valeur totale"].sum() >= objectif:
            break

    if "Produit" in df.columns:
        total_row = {col: df[col].sum() if pd.api.types.is_numeric_dtype(df[col]) else "" for col in df.columns}
        total_row["Produit"] = "TOTAL"
        df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)

    return df
