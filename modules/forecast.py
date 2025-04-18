
import pandas as pd

def run_forecast_simulation(df):
    df = df.copy()
    df["Quantité commandée"] = 0  # simulation simple
    df["Valeur ajoutée"] = 0.0
    df["Valeur totale"] = df["Valeur stock actuel"]
    df["Stock total après commande"] = df["Stock"]
    if "Produit" in df.columns:
        total_row = {
            col: df[col].sum() if df[col].dtype.kind in "if" else "" for col in df.columns
        }
        total_row["Produit"] = "TOTAL"
        df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)
    return df

def run_target_stock_sim(df, objectif):
    df = df.copy()

    # Repartir de zéro
    df["Quantité commandée"] = 0
    df["Valeur ajoutée"] = 0.0
    df["Valeur totale"] = df["Valeur stock actuel"]
    df["Stock total après commande"] = df["Stock"]

    conditionnement = df["Conditionnement"].replace(0, 1)
    df["Tarif d’achat"].fillna(0, inplace=True)

    while df["Valeur totale"].sum() < objectif:
        produit_eligible = df.copy()
        produit_eligible["Progression"] = (produit_eligible["Tarif d’achat"] * conditionnement)
        produit_eligible = produit_eligible[produit_eligible["Progression"] > 0]

        if produit_eligible.empty:
            break

        idx = produit_eligible["Valeur stock actuel"].idxmin()
        df.at[idx, "Quantité commandée"] += conditionnement[idx]
        df.at[idx, "Valeur ajoutée"] = df.at[idx, "Quantité commandée"] * df.at[idx, "Tarif d’achat"]
        df.at[idx, "Valeur totale"] = df.at[idx, "Valeur stock actuel"] + df.at[idx, "Valeur ajoutée"]
        df.at[idx, "Stock total après commande"] = df.at[idx, "Stock"] + df.at[idx, "Quantité commandée"]

        if (df["Valeur totale"].sum() - objectif) > df["Tarif d’achat"].max() * df["Conditionnement"].max():
            break

    if "Produit" in df.columns:
        total_row = {
            col: df[col].sum() if df[col].dtype.kind in "if" else "" for col in df.columns
        }
        total_row["Produit"] = "TOTAL"
        df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)

    return df
