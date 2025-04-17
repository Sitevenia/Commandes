# modules/rotation.py
import pandas as pd

def detect_low_rotation_products(df, threshold=10, include_zero_stock=True):
    df = df.copy()
    week_cols = [col for col in df.columns if isinstance(col, int) and 1 <= col <= 52]

    df['Rotation'] = df[week_cols].sum(axis=1) / len(week_cols)

    if not include_zero_stock:
        df = df[df['Quantité mini'] > 0]

    return df[df['Rotation'] < threshold][['Produit', 'Rotation', 'Quantité mini', 'Stock']]
