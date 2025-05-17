import streamlit as st
import pandas as pd
import os
from datetime import datetime

# Configuration de la page
st.set_page_config(page_title="Logistique - Suivi", page_icon="🚚")
st.title("🚚 Suivi Logistique")

# Chemins des fichiers Excel
ARRET_LIGNE_FILE = r"C:/Users/White/Desktop/clerprem-project/data/Suivi_Arret_Ligne.xlsx"
SUIVI_FOURNISSEUR_FILE = r"C:/Users/White/Desktop/clerprem-project/data/Suivi_Fournisseur.xlsx"

# Fonction pour initialiser un fichier Excel s'il n'existe pas
def init_excel_file(file_path, columns):
    if not os.path.exists(file_path):
        df = pd.DataFrame(columns=columns)
        df.to_excel(file_path, index=False)
        st.success(f"Fichier créé : {file_path}")
    elif not os.access(file_path, os.W_OK):
        st.error(f"Impossible d'écrire dans le fichier : {file_path}. Vérifiez les permissions.")
        st.stop()

try:
    # Colonnes pour le suivi arrêt ligne
    arret_ligne_columns = [
        'DATE', 'Département', 'Projet', 'Ligne', 'Nbr heures',
        'Composants', 'Cause', 'Actions', 'Status'
    ]
    
    # Colonnes pour le suivi fournisseur
    suivi_fournisseur_columns = [
        'Fournisseur', 'Status Fournisseur', 'Risque Arrêt', 'Date Arrêt',
        'Composants', 'Cause', 'Projet Impacté'
    ]
    
    # Création des fichiers s'ils n'existent pas
    init_excel_file(ARRET_LIGNE_FILE, arret_ligne_columns)
    init_excel_file(SUIVI_FOURNISSEUR_FILE, suivi_fournisseur_columns)
    
    # Onglets pour naviguer entre les formulaires
    tab1, tab2 = st.tabs(["Suivi Arrêt Ligne", "Suivi Fournisseur"])
    
    # Formulaire pour le suivi arrêt ligne
    with tab1:
        st.header("Nouvel Arrêt de Ligne")
        with st.form("form_arret_ligne"):
            col1, col2 = st.columns(2)
            with col1:
                date = st.date_input("Date*", value=datetime.now())
                departement = st.text_input("Département*")
                projet = st.text_input("Projet*")
                ligne = st.text_input("Ligne*")
            with col2:
                nbr_heures = st.number_input("Nombre d'heures*", min_value=0.0, step=0.5)
                status = st.selectbox("Statut*", ["En cours", "Résolu", "En attente"])
            
            composants = st.text_area("Composants concernés*")
            cause = st.text_area("Cause de l'arrêt*")
            actions = st.text_area("Actions entreprises ou prévues")
            
            submit_arret = st.form_submit_button("Enregistrer l'arrêt de ligne")
            
            if submit_arret:
                if not all([departement, projet, ligne, composants, cause]):
                    st.error("Veuillez remplir tous les champs obligatoires (*)")
                else:
                    try:
                        new_data = {
                            'DATE': date,
                            'Département': departement,
                            'Projet': projet,
                            'Ligne': ligne,
                            'Nbr heures': nbr_heures,
                            'Composants': composants,
                            'Cause': cause,
                            'Actions': actions,
                            'Status': status
                        }
                        
                        # Lire les données existantes
                        if os.path.exists(ARRET_LIGNE_FILE):
                            df = pd.read_excel(ARRET_LIGNE_FILE)
                        else:
                            df = pd.DataFrame(columns=arret_ligne_columns)
                        
                        # Ajouter la nouvelle entrée
                        df = pd.concat([df, pd.DataFrame([new_data])], ignore_index=True)
                        
                        # Sauvegarder
                        df.to_excel(ARRET_LIGNE_FILE, index=False)
                        st.success("Arrêt de ligne enregistré avec succès!")
                        
                    except Exception as e:
                        st.error(f"Erreur lors de l'enregistrement : {e}")
    
    # Formulaire pour le suivi fournisseur
    with tab2:
        st.header("Nouveau Suivi Fournisseur")
        with st.form("form_fournisseur"):
            col1, col2 = st.columns(2)
            with col1:
                fournisseur = st.text_input("Fournisseur*")
                status_fournisseur = st.selectbox("Status Fournisseur*", 
                                                 ["En règle", "En alerte", "En retard", "En litige"])
                risque_arret = st.selectbox("Risque d'arrêt*", 
                                           ["Aucun", "Faible", "Moyen", "Élevé", "Arrêt en cours"])
            with col2:
                date_arret = st.date_input("Date d'arrêt prévue/effective*", value=datetime.now())
                projet_impacte = st.text_input("Projet impacté*")
            
            composants = st.text_area("Composants concernés*")
            cause = st.text_area("Cause du problème*")
            
            submit_fournisseur = st.form_submit_button("Enregistrer le suivi fournisseur")
            
            if submit_fournisseur:
                if not all([fournisseur, composants, cause, projet_impacte]):
                    st.error("Veuillez remplir tous les champs obligatoires (*)")
                else:
                    try:
                        new_data = {
                            'Fournisseur': fournisseur,
                            'Status Fournisseur': status_fournisseur,
                            'Risque Arrêt': risque_arret,
                            'Date Arrêt': date_arret,
                            'Composants': composants,
                            'Cause': cause,
                            'Projet Impacté': projet_impacte
                        }
                        
                        # Lire les données existantes
                        if os.path.exists(SUIVI_FOURNISSEUR_FILE):
                            df = pd.read_excel(SUIVI_FOURNISSEUR_FILE)
                        else:
                            df = pd.DataFrame(columns=suivi_fournisseur_columns)
                        
                        # Ajouter la nouvelle entrée
                        df = pd.concat([df, pd.DataFrame([new_data])], ignore_index=True)
                        
                        # Sauvegarder
                        df.to_excel(SUIVI_FOURNISSEUR_FILE, index=False)
                        st.success("Suivi fournisseur enregistré avec succès!")
                        
                    except Exception as e:
                        st.error(f"Erreur lors de l'enregistrement : {e}")
    
    # Affichage des données existantes dans la barre latérale
    st.sidebar.title("Afficher les données")
    show_arret_ligne = st.sidebar.checkbox("Afficher les arrêts de ligne")
    show_fournisseur = st.sidebar.checkbox("Afficher le suivi fournisseurs")
    
    if show_arret_ligne and os.path.exists(ARRET_LIGNE_FILE):
        st.sidebar.subheader("Derniers arrêts de ligne")
        df_arret = pd.read_excel(ARRET_LIGNE_FILE)
        st.sidebar.dataframe(df_arret.tail(5))
    
    if show_fournisseur and os.path.exists(SUIVI_FOURNISSEUR_FILE):
        st.sidebar.subheader("Derniers suivis fournisseurs")
        df_fourn = pd.read_excel(SUIVI_FOURNISSEUR_FILE)
        st.sidebar.dataframe(df_fourn.tail(5))

except Exception as e:
    st.error(f"Une erreur est survenue : {e}")
    st.stop()
