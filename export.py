# modules/export.py
import os
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

def export_order_pdfs(df):
    df = df[df['Commande simulée'] > 0]
    grouped = df.groupby('Fournisseur')

    for fournisseur, group in grouped:
        filename = f"output/pdf/bon_commande_{fournisseur.replace(' ', '_')}.pdf"
        c = canvas.Canvas(filename, pagesize=A4)
        c.setFont("Helvetica", 12)
        c.drawString(50, 800, f"Bon de commande - Fournisseur : {fournisseur}")
        y = 770
        for _, row in group.iterrows():
            line = f"{row['Produit']} - Qté : {row['Commande simulée']}"
            c.drawString(50, y, line)
            y -= 20
            if y < 50:
                c.showPage()
                c.setFont("Helvetica", 12)
                y = 800
        c.save()

def export_low_rotation_list(df):
    output_path = "output/excel/faible_rotation.xlsx"
    df.to_excel(output_path, index=False)
