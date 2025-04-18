
import pandas as pd

def run_forecast_simulation(df):
    df = df.copy()

    # Sécurisation des colonnes indispensables
    if "Tarif d’achat" not in df.columns:
        df["Tarif d’achat"] = 0.0
    if "Valeur stock actuel" not in df.columns:
        df["Valeur stock actuel"] = df.get("Stock", 0) * df.get("Tarif d’achat", 0)

    df["Quantité commandée"] = 0
    df["Valeur ajoutée"] = 0.0
    df["Valeur totale"] = df["Valeur stock actuel"]
    df["Stock total après commande"] = df.get("Stock", 0)

    if "Produit" in df.columns:
        total_row = {
            col: df[col].sum() if pd.api.types.is_numeric_dtype(df[col]) else "" for col in df.columns
        }
        total_row["Produit"] = "TOTAL"
        df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)

    return df

def run_target_stock_sim(df, objectif):
    df = df.copy()

    # Sécurisation des colonnes indispensables
    if "Tarif d’achat" not in df.columns:
        df["Tarif d’achat"] = 0.0
    if "Valeur stock actuel" not in df.columns:
        df["Valeur stock actuel"] = df.get("Stock", 0) * df.get("Tarif d’achat", 0)

    df["Quantité commandée"] = 0
    df["Valeur ajoutée"] = 0.0
    df["Valeur totale"] = df["Valeur stock actuel"]
    df["Stock total après commande"] = df.get("Stock", 0)

    df["Conditionnement"] = df.get("Conditionnement", 1).replace(0, 1)
    df["Tarif d’achat"].fillna(0, inplace=True)

    while df["Valeur totale"].sum() < objectif:
        produit_eligible = df.copy()
        produit_eligible["Progression"] = produit_eligible["Tarif d’achat"] * df["Conditionnement"]
        produit_eligible = produit_eligible[produit_eligible["Progression"] > 0]

        if produit_eligible.empty:
            break

        idx = produit_eligible["Valeur stock actuel"].idxmin()
        df.at[idx, "Quantité commandée"] += df.at[idx, "Conditionnement"]
        df.at[idx, "Valeur ajoutée"] = df.at[idx, "Quantité commandée"] * df.at[idx, "Tarif d’achat"]
        df.at[idx, "Valeur totale"] = df.at[idx, "Valeur stock actuel"] + df.at[idx, "Valeur ajoutée"]
        df.at[idx, "Stock total après commande"] = df.at[idx, "Stock"] + df.at[idx, "Quantité commandée"]

        if df["Valeur totale"].sum() >= objectif:
            break

    if "Produit" in df.columns:
        total_row = {
            col: df[col].sum() if pd.api.types.is_numeric_dtype(df[col]) else "" for col in df.columns
        }
        total_row["Produit"] = "TOTAL"
        df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)

    return df
