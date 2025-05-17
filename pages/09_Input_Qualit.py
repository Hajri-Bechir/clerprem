import streamlit as st
import pandas as pd
import os
from datetime import datetime

# Configuration de la page
st.set_page_config(page_title="Qualité - Réclamations", page_icon="📝")
st.title("📝 Gestion des Réclamations Qualité")

# Chemins des fichiers Excel
CLIENTS_FILE = r"C:/Users/White/Desktop/clerprem-project/data/Reclamations_Clients.xlsx"
FOURNISSEURS_FILE = r"C:/Users/White/Desktop/clerprem-project/data/Reclamations_Fournisseurs.xlsx"

# Fonction pour initialiser un fichier Excel s'il n'existe pas
def init_excel_file(file_path, columns):
    if not os.path.exists(file_path):
        df = pd.DataFrame(columns=columns)
        df.to_excel(file_path, index=False)
        st.success(f"Fichier créé : {file_path}")
    elif not os.access(file_path, os.W_OK):
        st.error(f"Impossible d'écrire dans le fichier : {file_path}. Vérifiez les permissions.")
        st.stop()

# Initialisation des fichiers
try:
    # Colonnes pour les réclamations clients
    clients_columns = ['Clients', 'Projets', 'Semaine', 'Date', 'Descriptions', 
                      'Tri chez Le client', 'Causes', 'Actions', 'Status']
    
    # Colonnes pour les réclamations fournisseurs
    fournisseurs_columns = ['Fournisseurs', 'Références', 'Semaine', 'Date', 'Descriptions', 
                           'Trie', 'Causes', 'Actions', 'Status']
    
    # Création des fichiers s'ils n'existent pas
    init_excel_file(CLIENTS_FILE, clients_columns)
    init_excel_file(FOURNISSEURS_FILE, fournisseurs_columns)
    
    # Onglets pour naviguer entre les formulaires
    tab1, tab2 = st.tabs(["Réclamations Clients", "Réclamations Fournisseurs"])
    
    # Formulaire pour les réclamations clients
    with tab1:
        st.header("Nouvelle Réclamation Client")
        with st.form("form_clients"):
            col1, col2 = st.columns(2)
            with col1:
                client = st.text_input("Client*")
                projet = st.text_input("Projet*")
                semaine = st.number_input("Semaine*", min_value=1, max_value=53, value=datetime.now().isocalendar()[1])
                date = st.date_input("Date*", value=datetime.now())
            
            description = st.text_area("Description de la réclamation*")
            tri_client = st.selectbox("Tri chez le client*", ["Oui", "Non", "En cours"])
            causes = st.text_area("Causes identifiées")
            actions = st.text_area("Actions correctives")
            status = st.selectbox("Statut*", ["Nouveau", "En cours", "Résolu", "Fermé"])
            
            submit_client = st.form_submit_button("Enregistrer la réclamation client")
            
            if submit_client:
                if not all([client, projet, description]):
                    st.error("Veuillez remplir tous les champs obligatoires (*)")
                else:
                    try:
                        new_data = {
                            'Clients': client,
                            'Projets': projet,
                            'Semaine': semaine,
                            'Date': date,
                            'Descriptions': description,
                            'Tri chez Le client': tri_client,
                            'Causes': causes,
                            'Actions': actions,
                            'Status': status
                        }
                        
                        # Lire les données existantes
                        if os.path.exists(CLIENTS_FILE):
                            df = pd.read_excel(CLIENTS_FILE)
                        else:
                            df = pd.DataFrame(columns=clients_columns)
                        
                        # Ajouter la nouvelle entrée
                        df = pd.concat([df, pd.DataFrame([new_data])], ignore_index=True)
                        
                        # Sauvegarder
                        df.to_excel(CLIENTS_FILE, index=False)
                        st.success("Réclamation client enregistrée avec succès!")
                        
                    except Exception as e:
                        st.error(f"Erreur lors de l'enregistrement : {e}")
    
    # Formulaire pour les réclamations fournisseurs
    with tab2:
        st.header("Nouvelle Réclamation Fournisseur")
        with st.form("form_fournisseurs"):
            col1, col2 = st.columns(2)
            with col1:
                fournisseur = st.text_input("Fournisseur*")
                reference = st.text_input("Référence*")
                semaine = st.number_input("Semaine* ", min_value=1, max_value=53, value=datetime.now().isocalendar()[1])
                date = st.date_input("Date* ", value=datetime.now())
            
            description = st.text_area("Description de la réclamation* ", key="desc_fourn")
            trie = st.selectbox("Tri* ", ["Oui", "Non", "En attente"], key="trie_fourn")
            causes = st.text_area("Causes identifiées", key="causes_fourn")
            actions = st.text_area("Actions correctives", key="actions_fourn")
            status = st.selectbox("Statut* ", ["Nouveau", "En cours", "Résolu", "Fermé"], key="status_fourn")
            
            submit_fournisseur = st.form_submit_button("Enregistrer la réclamation fournisseur")
            
            if submit_fournisseur:
                if not all([fournisseur, reference, description]):
                    st.error("Veuillez remplir tous les champs obligatoires (*)")
                else:
                    try:
                        new_data = {
                            'Fournisseurs': fournisseur,
                            'Références': reference,
                            'Semaine': semaine,
                            'Date': date,
                            'Descriptions': description,
                            'Trie': trie,
                            'Causes': causes,
                            'Actions': actions,
                            'Status': status
                        }
                        
                        # Lire les données existantes
                        if os.path.exists(FOURNISSEURS_FILE):
                            df = pd.read_excel(FOURNISSEURS_FILE)
                        else:
                            df = pd.DataFrame(columns=fournisseurs_columns)
                        
                        # Ajouter la nouvelle entrée
                        df = pd.concat([df, pd.DataFrame([new_data])], ignore_index=True)
                        
                        # Sauvegarder
                        df.to_excel(FOURNISSEURS_FILE, index=False)
                        st.success("Réclamation fournisseur enregistrée avec succès!")
                        
                    except Exception as e:
                        st.error(f"Erreur lors de l'enregistrement : {e}")
    
    # Affichage des données existantes dans la barre latérale
    st.sidebar.title("Afficher les données")
    show_clients = st.sidebar.checkbox("Afficher les réclamations clients")
    show_fournisseurs = st.sidebar.checkbox("Afficher les réclamations fournisseurs")
    
    if show_clients and os.path.exists(CLIENTS_FILE):
        st.sidebar.subheader("Dernières réclamations clients")
        df_clients = pd.read_excel(CLIENTS_FILE)
        st.sidebar.dataframe(df_clients.tail(5))
    
    if show_fournisseurs and os.path.exists(FOURNISSEURS_FILE):
        st.sidebar.subheader("Dernières réclamations fournisseurs")
        df_fourn = pd.read_excel(FOURNISSEURS_FILE)
        st.sidebar.dataframe(df_fourn.tail(5))

except Exception as e:
    st.error(f"Une erreur est survenue : {e}")
    st.stop()
