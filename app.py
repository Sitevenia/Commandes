import streamlit as st
import pandas as pd
import numpy as np
import io

def repartir_et_ajuster(total_qte, ventes_n1_mois, conditionnement):
    """Répartit une quantité totale selon la saisonnalité et ajuste aux conditionnements."""
    if total_qte <= 0 or conditionnement is None or conditionnement <= 0:
        return [0] * len(ventes_n1_mois)

    total_ventes = sum(ventes_n1_mois)
    if total_ventes == 0:
        repartition = [1] * len(ventes_n1_mois)
        total_ventes = len(ventes_n1_mois)
    else:
        repartition = ventes_n1_mois

    proportions = [v / total_ventes for v in repartition]
    qtes_mensuelles = [round(total_qte * p) for p in proportions]

    # Ajuster chaque mois au multiple de conditionnement
    qtes_conditionnees = [int(round(q / conditionnement)) * conditionnement for q in qtes_mensuelles]

    # Réajuster si trop ou pas assez
    ecart = sum(qtes_conditionnees) - total_qte
    while ecart != 0:
        for i in range(len(qtes_conditionnees)):
            if ecart > 0 and qtes_conditionnees[i] >= conditionnement:
                qtes_conditionnees[i] -= conditionnement
                ecart -= conditionnement
            elif ecart < 0:
                qtes_conditionnees[i] += conditionnement
                ecart += conditionnement
            if ecart == 0:
                break

    return qtes_conditionnees

st.set_page_config(page_title="Forecast App", layout="wide")
st.title("📦 Application de Prévision des Commandes")

def safe_int(x):
    try:
        if 
