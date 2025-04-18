
import pandas as pd

def run_forecast_simulation(df):
    df = df.copy()

    if "Valeur stock actuel" not in df.columns:
        df["Valeur stock actuel"] = df["Stock"] * df["Tarif d’achat"]

    df["Quantité commandée"] = 0
    df["Valeur ajoutée"] = 0.0
    df["Valeur totale"] = df["Valeur stock actuel"]
    df["Stock total après commande"] = df["Stock"]

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

    produit_eligibles = df[df["Tarif d’achat"] > 0].copy()
    produit_eligibles["Score"] = (
        produit_eligibles["Valeur stock actuel"] + 1
    ) / (produit_eligibles["Tarif d’achat"] + 1)

    produit_eligibles = produit_eligibles.sort_values(by="Score")

    i = 0
    while df["Valeur totale"].sum() < objectif and i < len(produit_eligibles):
        idx = produit_eligibles.index[i]

        df.at[idx, "Quantité commandée"] += conditionnement[idx]
        df.at[idx, "Valeur ajoutée"] = df.at[idx, "Quantité commandée"] * df.at[idx, "Tarif d’achat"]
        df.at[idx, "Valeur totale"] = df.at[idx, "Valeur stock actuel"] + df.at[idx, "Valeur ajoutée"]
        df.at[idx, "Stock total après commande"] = df.at[idx, "Stock"] + df.at[idx, "Quantité commandée"]

        # On repart de 0 une fois la boucle terminée pour mieux lisser l’augmentation
        i = (i + 1) % len(produit_eligibles)

        if df["Valeur totale"].sum() >= objectif:
            break

    if "Produit" in df.columns:
        total_row = {col: df[col].sum() if pd.api.types.is_numeric_dtype(df[col]) else "" for col in df.columns}
        total_row["Produit"] = "TOTAL"
        df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)

    return df
