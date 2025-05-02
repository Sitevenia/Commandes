import streamlit as st
import pandas as pd
import numpy as np
import io

def calculer_quantite_a_commander(df, semaine_columns, montant_minimum, duree_semaines):
    """
    Calcule la quantité à commander pour chaque produit en fonction des ventes passées,
    du stock actuel, du conditionnement et d'un montant minimum de commande.

    Args:
        df (pd.DataFrame): DataFrame contenant les données des produits (doit inclure
                           'Stock', 'Conditionnement', 'Tarif d'achat' et les colonnes semaine).
        semaine_columns (list): Liste des noms des colonnes représentant les ventes hebdomadaires.
        montant_minimum (float): Montant minimum de commande pour le fournisseur.
        duree_semaines (int): Nombre de semaines de ventes à couvrir par la commande.

    Returns:
        tuple: Contient les éléments suivants:
               - list: Quantités à commander pour chaque produit.
               - pd.Series: Ventes totales N-1 pour chaque produit.
               - pd.Series: Ventes des 12 semaines N-1 équivalentes pour chaque produit.
               - pd.Series: Ventes des 12 dernières semaines pour chaque produit.
               - float: Montant total de la commande calculée.
        Retourne None en cas d'erreur.
    """
    try:
        # --- Validation des Entrées ---
        if not isinstance(df, pd.DataFrame) or df.empty:
            st.error("Le DataFrame d'entrée est vide ou invalide.")
            return None
        required_cols = ["Stock", "Conditionnement", "Tarif d'achat"] + semaine_columns
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            st.error(f"Colonnes manquantes dans le DataFrame: {', '.join(missing_cols)}")
            return None
        if not semaine_columns:
            st.error("La liste des colonnes de semaines de vente est vide.")
            return None
        # Assurer que les colonnes nécessaires sont numériques et gérer les NaN/Infs
        for col in required_cols:
             # Remplacer Inf par NaN puis remplir les NaN par 0
            df[col] = pd.to_numeric(df[col], errors='coerce').replace([np.inf, -np.inf], np.nan).fillna(0)


        # --- Calculs des Ventes Moyennes ---
        # S'assurer qu'il y a assez de colonnes pour les calculs N-1
        if len(semaine_columns) < 64:
            st.warning("Pas assez de données historiques (< 64 semaines) pour tous les calculs N-1. Certains calculs N-1 seront mis à zéro.")
            # Mettre à zéro les calculs N-1 s'il n'y a pas assez de données
            ventes_12_semaines_N1 = pd.Series(0, index=df.index)
            ventes_12_semaines_N1_suivantes = pd.Series(0, index=df.index)
            # Calculer les ventes N-1 totales avec les colonnes disponibles
            ventes_N1 = df[semaine_columns].sum(axis=1)
        else:
             # Calculer la moyenne des ventes sur la totalité des colonnes (Ventes N-1)
            ventes_N1 = df[semaine_columns].sum(axis=1)
            # Calculer la somme des 12 semaines identiques en N-1
            ventes_12_semaines_N1 = df[semaine_columns[-64:-52]].sum(axis=1)
            # Calculer la somme des 12 semaines suivantes en N-1
            ventes_12_semaines_N1_suivantes = df[semaine_columns[-52:-40]].sum(axis=1)

        # S'assurer qu'il y a assez de colonnes pour les 12 dernières semaines
        if len(semaine_columns) < 12:
             st.warning("Pas assez de données historiques (< 12 semaines) pour le calcul des 12 dernières semaines. Ce calcul sera basé sur les semaines disponibles.")
             ventes_12_dernieres_semaines = df[semaine_columns].sum(axis=1) # Somme sur toutes les semaines dispo si < 12
             nb_semaines_recentes = len(semaine_columns)
        else:
            # Calculer la somme des 12 dernières semaines
            ventes_12_dernieres_semaines = df[semaine_columns[-12:]].sum(axis=1)
            nb_semaines_recentes = 12


        # --- Calcul de la Quantité Pondérée ---
        # Gérer la division par zéro si nb_semaines_recentes ou 12 est zéro (ne devrait pas arriver ici mais par sécurité)
        avg_12_dernieres = ventes_12_dernieres_semaines / nb_semaines_recentes if nb_semaines_recentes > 0 else 0
        avg_12_N1 = ventes_12_semaines_N1 / 12 if len(semaine_columns) >= 64 else 0
        avg_12_N1_suivantes = ventes_12_semaines_N1_suivantes / 12 if len(semaine_columns) >= 64 else 0

        quantite_ponderee = (0.5 * avg_12_dernieres +
                             0.2 * avg_12_N1 +
                             0.3 * avg_12_N1_suivantes)

        # Calculer la quantité nécessaire pour couvrir les ventes pendant la durée spécifiée
        quantite_necessaire = quantite_ponderee * duree_semaines

        # Calculer la quantité à commander initiale
        quantite_a_commander_series = quantite_necessaire - df["Stock"]
        # Ne pas commander des quantités négatives
        quantite_a_commander_series = quantite_a_commander_series.apply(lambda x: max(0, x))

        # --- Ajustements Basés sur les Règles ---
        conditionnement = df["Conditionnement"]
        stock_actuel = df["Stock"]
        tarif_achat = df["Tarif d'achat"]

        # Convertir en liste pour modifications potentielles
        quantite_a_commander = quantite_a_commander_series.tolist()

        # Ajuster les quantités aux conditionnements (arrondi supérieur)
        for i in range(len(quantite_a_commander)):
            cond = conditionnement.iloc[i]
            q = quantite_a_commander[i]
             # Vérifier si conditionnement est > 0 pour éviter la division par zéro
            if q > 0 and cond > 0:
                 # np.ceil retourne un float, convertir en int
                quantite_a_commander[i] = int(np.ceil(q / cond) * cond)
            elif q > 0 and cond <= 0:
                 # Si conditionnement est 0 ou négatif, on ne peut pas arrondir. Garder q ou mettre 0 ?
                 # On pourrait logguer une alerte ou mettre 0. Mettons 0 pour la sécurité.
                 st.warning(f"Conditionnement invalide (<=0) pour le produit index {i}. Quantité mise à 0.")
                 quantite_a_commander[i] = 0
            else: # q <= 0
                quantite_a_commander[i] = 0 # Assurer que c'est bien 0

        # Règle 1: Vendu >= 2 fois (12 dernières semaines) ET Stock <= 1 => Commander au moins 1 conditionnement
        if nb_semaines_recentes > 0: # Appliquer seulement si on a des données récentes
            for i in range(len(quantite_a_commander)):
                cond = conditionnement.iloc[i]
                # Compter le nombre de semaines avec ventes > 0 dans les dernières semaines
                ventes_recentes_count = (df[semaine_columns[-nb_semaines_recentes:]].iloc[i] > 0).sum()
                if ventes_recentes_count >= 2 and stock_actuel.iloc[i] <= 1 and cond > 0:
                    quantite_a_commander[i] = max(quantite_a_commander[i], cond)

        # Règle 2: Ventes N-1 < 6 ET Ventes 12 dernières semaines < 2 => Ne pas commander (mettre à 0)
        for i in range(len(quantite_a_commander)):
            # Utiliser les ventes totales N-1 calculées précédemment
            ventes_tot_n1 = ventes_N1.iloc[i]
            # Utiliser les ventes des 12 (ou moins) dernières semaines calculées précédemment
            ventes_recentes_sum = ventes_12_dernieres_semaines.iloc[i]
            if ventes_tot_n1 < 6 and ventes_recentes_sum < 2:
                quantite_a_commander[i] = 0

        # --- Ajustement pour Montant Minimum ---
        # Calculer le montant total avant ajustement minimum
        montant_total_actuel = sum(q * p for q, p in zip(quantite_a_commander, tarif_achat))

        if montant_minimum > 0 and montant_total_actuel < montant_minimum:
            # Créer une liste d'indices des produits ayant qte > 0 *initialement* ou après règles 1/2
            # pour prioriser l'augmentation de ceux-là.
            indices_commandes = [i for i, q in enumerate(quantite_a_commander) if q > 0]

            # Trier potentiellement les indices (ex: par prix décroissant pour atteindre le min plus vite?) - Optionnel
            # indices_commandes.sort(key=lambda i: tarif_achat.iloc[i] * conditionnement.iloc[i], reverse=True)

            # Boucle d'ajustement
            idx_pointer = 0
            while montant_total_actuel < montant_minimum:
                if not indices_commandes: # S'il n'y a aucun produit commandé initialement, on ne peut pas augmenter
                    st.warning(f"Impossible d'atteindre le montant minimum de {montant_minimum:.2f}€ car aucune quantité initiale n'était à commander. Montant actuel : {montant_total_actuel:.2f}€")
                    break # Sortir de la boucle while

                # Sélectionner l'indice à augmenter (cycle à travers les produits déjà commandés)
                current_idx = indices_commandes[idx_pointer % len(indices_commandes)]

                cond = conditionnement.iloc[current_idx]
                prix = tarif_achat.iloc[current_idx]

                # Vérifier si cond et prix sont valides pour éviter boucle infinie
                if cond > 0 and prix > 0 :
                    # Augmenter la quantité d'un conditionnement
                    quantite_a_commander[current_idx] += cond
                    # Mettre à jour le montant total
                    montant_total_actuel += cond * prix
                elif cond <= 0:
                     st.warning(f"Conditionnement invalide (<=0) rencontré lors de l'ajustement du montant minimum pour produit index {current_idx}. Ce produit est ignoré.")
                     # Retirer cet index de la liste pour ne pas retenter
                     indices_commandes.pop(idx_pointer % len(indices_commandes))
                     if not indices_commandes: continue # Vérifier si la liste est vide après suppression
                     idx_pointer -=1 # Ajuster le pointeur car un élément a été retiré avant lui
                # else: prix <= 0 -> l'augmentation n'aidera pas, on pourrait l'ignorer aussi

                # Passer à l'indice suivant pour la prochaine itération
                idx_pointer += 1

                # Sécurité: éviter une boucle infinie si l'augmentation est impossible (ex: tous les cond/prix sont 0)
                if idx_pointer > len(quantite_a_commander) * 5 and montant_total_actuel < montant_minimum : # Heuristique
                     st.error("Impossible d'atteindre le montant minimum après de nombreuses tentatives. Vérifiez les conditionnements et tarifs.")
                     break

        # Recalculer le montant final après l'ajustement potentiel
        montant_final = sum(q * p for q, p in zip(quantite_a_commander, tarif_achat))

        # Retourner les valeurs sous forme de tuple
        return (quantite_a_commander,
                ventes_N1,
                ventes_12_semaines_N1,
                ventes_12_dernieres_semaines,
                montant_final)

    except KeyError as e:
        st.error(f"Erreur de clé: Colonne '{e}' introuvable dans le DataFrame.")
        return None
    except ValueError as e:
         st.error(f"Erreur de valeur: Problème avec les données numériques - {e}")
         return None
    except Exception as e:
        st.error(f"Erreur inattendue dans le calcul des quantités à commander: {e}")
        import traceback
        st.error(traceback.format_exc()) # Affiche la trace complète pour le débogage
        return None

# --- Configuration de la page Streamlit ---
st.set_page_config(page_title="Forecast App", layout="wide")
st.title("📦 Application de Prévision des Commandes")

# --- Chargement du fichier ---
uploaded_file = st.file_uploader("📁 Charger le fichier Excel principal (ventes hebdo, stock, conditionnement...)", type=["xlsx", "xls"])

if uploaded_file:
    try:
        # Lire le fichier Excel, en spécifiant la ligne d'en-tête (0-based index, donc ligne 8 = header=7)
        df_full = pd.read_excel(uploaded_file, sheet_name="Tableau final", header=7)
        st.success("✅ Fichier principal chargé avec succès.")

        # --- Nettoyage Initial et Filtrage ---
        # Garder une copie originale si besoin plus tard
        # df_original = df_full.copy()

        # Filtrer les lignes invalides
        df = df_full[
            (df_full["Fournisseur"].notna()) &
            (df_full["Fournisseur"] != "") &
            (df_full["Fournisseur"] != "#FILTER") & # Exclure explicitement '#FILTER'
            (df_full["AF_RefFourniss"].notna()) &
            (df_full["AF_RefFourniss"] != "")
        ].copy() # Utiliser .copy() pour éviter SettingWithCopyWarning

        if df.empty:
             st.warning("Aucune ligne valide trouvée après le filtrage initial (Fournisseur/AF_RefFourniss renseignés et non '#FILTER').")

        # Extraire la liste des fournisseurs uniques et triés
        fournisseurs = sorted(df["Fournisseur"].unique().tolist())

        # --- Sélection Utilisateur ---
        selected_fournisseurs = st.multiselect(
            "👤 Sélectionnez le(s) fournisseur(s) pour le calcul",
            options=fournisseurs,
            default=[] # Par défaut, aucun fournisseur n'est sélectionné
        )

        # Filtrer les données en fonction des fournisseurs sélectionnés
        if selected_fournisseurs:
            df_filtered = df[df["Fournisseur"].isin(selected_fournisseurs)].copy() # Utiliser .copy()
        else:
            # Créer un DataFrame vide avec les bonnes colonnes si aucun fournisseur n'est sélectionné
            df_filtered = pd.DataFrame(columns=df.columns)

        # --- Identification et Nettoyage des Colonnes de Vente ---
        # Utiliser la colonne 13 (index 12) comme point de départ pour les semaines
        # Attention: les indices de colonnes peuvent changer si le fichier change. Nommer les colonnes est plus robuste.
        # Ici, on se base sur l'index comme demandé.
        start_col_index = 12 # Index de la colonne "N" (13ème colonne)

        if len(df_filtered.columns) > start_col_index:
            # Sélectionner toutes les colonnes à partir de l'index de départ
            potential_week_cols = df_filtered.columns[start_col_index:].tolist()

            # Identifier les colonnes qui semblent être des semaines de vente (numériques)
            # Exclure explicitement les colonnes connues non liées aux ventes hebdomadaires
            exclude_columns = ["Tarif d'achat", "Conditionnement", "Stock", "Total", "Stock à terme",
                               "Ventes N-1", "Ventes 12 semaines identiques N-1", "Ventes 12 dernières semaines",
                               "Quantité à commander" # Exclure aussi les colonnes qu'on va créer
                              ]

            # Garder uniquement les colonnes numériques potentielles non exclues
            semaine_columns = []
            for col in potential_week_cols:
                 # Vérifier si la colonne est potentiellement numérique (rapide check) et non exclue
                 # Le nettoyage final se fera dans la fonction de calcul
                 if col not in exclude_columns and pd.api.types.is_numeric_dtype(df_filtered[col].dtype):
                      semaine_columns.append(col)
                 # On pourrait ajouter un check sur le nom de colonne s'il suit un pattern (ex: 'S')
                 # elif col.startswith('S') and col not in exclude_columns: # Exemple
                 #    semaine_columns.append(col)


            st.info(f"Colonnes identifiées comme semaines de vente: {len(semaine_columns)} colonnes trouvées à partir de la colonne {start_col_index+1}.")
            # st.caption(f"Exemple de colonnes semaines: {semaine_columns[:5]} ... {semaine_columns[-5:]}") # Debug

            # S'assurer que les colonnes essentielles (Stock, Cond, Tarif) sont bien numériques
            essential_numeric_cols = ["Stock", "Conditionnement", "Tarif d'achat"]
            for col in essential_numeric_cols:
                 if col in df_filtered.columns:
                     df_filtered[col] = pd.to_numeric(df_filtered[col], errors='coerce').fillna(0)
                 else:
                     st.error(f"Colonne essentielle '{col}' manquante dans les données filtrées.")
                     # Arrêter le traitement si une colonne essentielle manque
                     st.stop()


        else:
            st.warning("Le fichier ne semble pas contenir de colonnes après l'index 12 pour les données de ventes.")
            semaine_columns = [] # Pas de colonnes de semaine

        # --- Paramètres de Calcul ---
        col1, col2 = st.columns(2)
        with col1:
            duree_semaines = st.number_input("⏳ Durée de couverture souhaitée (en semaines)", value=4, min_value=1, step=1, help="Nombre de semaines de ventes que la commande doit couvrir.")
        with col2:
            montant_minimum = st.number_input("💶 Montant minimum de commande (€)", value=0.0, min_value=0.0, step=50.0, format="%.2f", help="Montant minimum requis par le fournisseur pour passer commande. Laissez 0 si non applicable.")

        # --- Exécution du Calcul et Affichage ---
        if not df_filtered.empty and semaine_columns:
            st.info("🚀 Lancement du calcul des quantités à commander...")
            # Appeler la fonction de calcul
            result = calculer_quantite_a_commander(df_filtered, semaine_columns, montant_minimum, duree_semaines)

            if result is not None:
                st.success("✅ Calculs terminés.")
                # Dépaqueter les résultats
                (quantite_calculée, ventes_N1_calc, ventes_12_N1_calc,
                 ventes_12_last_calc, montant_total_calc) = result

                # Ajouter les résultats au DataFrame filtré
                # Utiliser .loc pour éviter SettingWithCopyWarning
                df_filtered.loc[:, "Quantité à commander"] = quantite_calculée
                df_filtered.loc[:, "Ventes N-1"] = ventes_N1_calc
                df_filtered.loc[:, "Ventes 12 semaines identiques N-1"] = ventes_12_N1_calc
                df_filtered.loc[:, "Ventes 12 dernières semaines"] = ventes_12_last_calc

                # Calculer les colonnes dérivées
                df_filtered.loc[:, "Total"] = df_filtered["Tarif d'achat"] * df_filtered["Quantité à commander"]
                df_filtered.loc[:, "Stock à terme"] = df_filtered["Stock"] + df_filtered["Quantité à commander"]

                # Afficher le montant total de la commande
                st.metric(label="💰 Montant total de la commande calculée", value=f"{montant_total_calc:.2f} €")
                if montant_minimum > 0 and montant_total_calc < montant_minimum:
                     st.warning(f"⚠️ Le montant total ({montant_total_calc:.2f}€) est inférieur au minimum requis ({montant_minimum:.2f}€) et n'a pas pu être ajusté (vérifiez conditionnements/prix ou produits commandés initialement).")
                elif montant_minimum > 0 and montant_total_calc >= montant_minimum and montant_total_calc != sum(df_filtered.loc[:, "Total"]) :
                     # Cette condition peut arriver si l'ajustement a eu lieu. On affiche le montant recalculé.
                      st.info(f"Le montant a été ajusté pour atteindre le minimum de {montant_minimum:.2f}€.")


                # --- Affichage du Tableau de Résultats ---
                st.subheader("📊 Résultats de la prévision de commande")

                # Vérifier si les colonnes nécessaires pour l'affichage existent
                required_display_columns = ["AF_RefFourniss", "Référence Article", "Désignation Article", "Stock"]
                missing_display_columns = [col for col in required_display_columns if col not in df_filtered.columns]

                if missing_display_columns:
                    st.error(f"❌ Colonnes manquantes pour l'affichage des résultats : {', '.join(missing_display_columns)}")
                else:
                    # Organiser l'ordre des colonnes pour un affichage clair
                    display_columns = [
                        "AF_RefFourniss", "Référence Article", "Désignation Article", "Stock",
                        "Ventes N-1", "Ventes 12 semaines identiques N-1", "Ventes 12 dernières semaines",
                        "Conditionnement", "Quantité à commander", "Stock à terme",
                        "Tarif d'achat", "Total"
                    ]
                    # Filtrer pour n'afficher que les colonnes existantes dans le DataFrame
                    display_columns = [col for col in display_columns if col in df_filtered.columns]

                    # Afficher le DataFrame avec les résultats
                    st.dataframe(df_filtered[display_columns].style.format({ # Appliquer un formatage
                        "Tarif d'achat": "{:.2f}€",
                        "Total": "{:.2f}€",
                        "Ventes N-1": "{:.0f}",
                        "Ventes 12 semaines identiques N-1": "{:.0f}",
                        "Ventes 12 dernières semaines": "{:.0f}",
                        "Stock": "{:.0f}",
                        "Conditionnement": "{:.0f}",
                        "Quantité à commander": "{:.0f}",
                        "Stock à terme": "{:.0f}"
                    }))

                    # --- Préparation et Export des Données ---
                    st.subheader("⬇️ Exportation des résultats")
                    # Filtrer les produits pour lesquels une quantité est à commander
                    df_export = df_filtered[df_filtered["Quantité à commander"] > 0].copy()

                    if not df_export.empty:
                        # Préparer le DataFrame pour l'export (sélectionner/ordonner colonnes + total)
                        export_columns = display_columns # Utiliser les mêmes colonnes que l'affichage
                        df_export_final = df_export[export_columns].copy()

                        # Ajouter une ligne de total
                        total_row_dict = {col: "" for col in export_columns} # Initialiser avec vide
                        total_row_dict["Désignation Article"] = "TOTAL COMMANDE" # Mettre le label dans une colonne texte
                        total_row_dict["Total"] = df_export_final["Total"].sum() # Calculer la somme de la colonne Total
                        # Convertir le dict en DataFrame
                        total_row_df = pd.DataFrame([total_row_dict])
                        # Concaténer avec les données
                        df_with_total = pd.concat([df_export_final, total_row_df], ignore_index=True)

                        # Créer le fichier Excel en mémoire
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine="openpyxl") as writer:
                            # Écrire sans l'index pandas, appliquer formatage si possible (plus complexe avec BytesIO)
                             df_with_total.to_excel(writer, sheet_name="Quantités_à_commander", index=False)
                            # On pourrait ajouter du formatage ici avec xlsxwriter ou openpyxl si nécessaire

                        output.seek(0) # Revenir au début du buffer

                        # Créer le nom du fichier dynamiquement
                        suppliers_str = "_".join(selected_fournisseurs).replace(" ", "_") # Créer une string des fournisseurs
                        filename = f"commande_{suppliers_str}_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx"

                        st.download_button(
                            label="📥 Télécharger le fichier de commande (Excel)",
                            data=output,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.info("ℹ️ Aucune quantité à commander pour les fournisseurs sélectionnés avec les paramètres actuels.")

            else:
                # Erreur gérée et affichée dans la fonction calculer_quantite_a_commander
                st.error("❌ Le calcul des quantités n'a pas pu aboutir. Vérifiez les messages d'erreur ci-dessus.")
        elif not selected_fournisseurs:
            st.warning("⚠️ Veuillez sélectionner au moins un fournisseur pour lancer le calcul.")
        elif not semaine_columns:
             st.warning("⚠️ Impossible de lancer le calcul car aucune colonne de ventes hebdomadaires n'a été identifiée ou le fichier est incomplet.")
        else: # df_filtered est vide mais des fournisseurs sont sélectionnés (rare)
             st.warning("⚠️ Aucun produit trouvé pour le(s) fournisseur(s) sélectionné(s) après le filtrage initial.")

    except FileNotFoundError:
         st.error("❌ Erreur : Le fichier spécifié n'a pas été trouvé. Vérifiez le chemin.")
    except ValueError as e:
         st.error(f"❌ Erreur de valeur lors de la lecture du fichier Excel. Vérifiez le format du fichier et la ligne d'en-tête spécifiée (header=7). Détails : {e}")
    except Exception as e:
        st.error(f"❌ Une erreur inattendue est survenue lors du chargement ou du traitement initial du fichier : {e}")
        import traceback
        st.error(traceback.format_exc()) # Utile pour le débogage

else:
    st.info("👋 Bienvenue ! Veuillez charger votre fichier Excel contenant les données de ventes et de stock pour commencer.")
