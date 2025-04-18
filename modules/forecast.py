import pandas as pd
import numpy as np

def run_forecast_simulation(df):
    df = df.copy()
    week_cols = [col for col in df.columns if isinstance(col, int) and 1 <= col <= 52]
    current_week = max([col for col in week_cols if col <= pd.Timestamp.today().week])

    last_12_weeks = [w for w in range(current_week - 11, current_week + 1) if w in week_cols]
    same_weeks_last_year = [w for w in last_12_weeks]

    df['Moyenne_12s'] = df[last_12_weeks].mean(axis=1)
    df['Moyenne_N-1'] = df[[f"{w}_N-1" for w in same_weeks_last_year if f"{w}_N-1" in df.columns]].mean(axis=1)

    df['Prévision'] = (df['Moyenne_12s'] + df['Moyenne_N-1']) / 2
    df['Prévision'] = df['Prévision'].fillna(0)

    df['Max_s'] = df[last_12_weeks].max(axis=1)
    df['Ratio_max_moy'] = df['Max_s'] / df['Moyenne_12s']
    df['Anomalie'] = df['Ratio_max_moy'] > 3

    for i, row in df.iterrows():
        if row['Anomalie']:
            if row['Prévision'] > 0:
                df.at[i, 'Prévision'] = row['Moyenne_N-1']

    ventes_rec = (df[last_12_weeks] > 0).sum(axis=1)
    df.loc[ventes_rec == 1, 'Prévision'] = 0

    df['Commande simulée'] = 0

    for i, row in df.iterrows():
        prevision = row['Prévision']
        stock = row['Stock']
        mini = row['Quantité mini']
        cond = row['Conditionnement']

        if mini == 0:
            if stock < 0:
                besoin = -stock
                qte = int(np.ceil(besoin / cond)) * cond
            else:
                qte = 0
        else:
            if stock < mini:
                besoin = mini - stock
                qte = int(np.ceil(besoin / cond)) * cond
            else:
                qte = 0

        df.at[i, 'Commande simulée'] = qte

    return df

def run_target_stock_simulation(df, objectif):
    df = run_forecast_simulation(df)
    df = df.sort_values(by='Rotation', ascending=False)
    total_value = (df['Commande simulée'] * df['Tarif achat']).sum()

    while total_value > objectif and df['Commande simulée'].sum() > 0:
        for i, row in df.iterrows():
            if row['Commande simulée'] >= row['Conditionnement']:
                df.at[i, 'Commande simulée'] -= row['Conditionnement']
        total_value = (df['Commande simulée'] * df['Tarif achat']).sum()

    return df
