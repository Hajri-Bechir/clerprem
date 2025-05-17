import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta

# Configuration de la page
st.set_page_config(page_title="Suivi Technique - Machines", page_icon="⚙️")
st.title("⚙️ Suivi Technique des Machines")

# Chemins des fichiers Excel
G11_FILE = r"C:/Users/White/Desktop/clerprem-project/data/Suivi_G11.xlsx"
G20_FILE = r"C:/Users/White/Desktop/clerprem-project/data/Suivi_G20.xlsx"
C12_FILE = r"C:/Users/White/Desktop/clerprem-project/data/Suivi_C12.xlsx"

# Jours de la semaine
JOURS_SEMAINE = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi"]

# Fonction pour initialiser un fichier Excel pour une machine
def init_machine_file(file_path, machine_name):
    if not os.path.exists(file_path):
        data = {
            'Jour': JOURS_SEMAINE + ['Total Minutes', 'Total Heures'],
            'Temps D\'ouverture': [1440] * 6 + [8640, 144],
            'Arrêt Machine (minutes)': [0] * 8,
            'Disponibilité Machine': ['100%'] * 6 + ['100%', ''],
            'Interventions': pd.Series([''] * 8, dtype='string')
        }
        df = pd.DataFrame(data)
        df.to_excel(file_path, index=False, sheet_name=machine_name)
        st.success(f"Fichier créé : {file_path}")
    elif not os.access(file_path, os.W_OK):
        st.error(f"Impossible d'écrire dans le fichier : {file_path}. Vérifiez les permissions.")
        st.stop()

# Fonction pour charger les données d'une machine
def load_machine_data(file_path):
    if os.path.exists(file_path):
        df = pd.read_excel(file_path)
        # S'assurer que la colonne Interventions est de type string
        if 'Interventions' in df.columns:
            df['Interventions'] = df['Interventions'].astype('string').fillna('')
        return df
    return None

# Fonction pour sauvegarder les données d'une machine
def save_machine_data(file_path, df, machine_name):
    # Calculer les totaux
    total_minutes = df.loc[df['Jour'].isin(JOURS_SEMAINE), 'Temps D\'ouverture'].sum()
    total_arret = df.loc[df['Jour'].isin(JOURS_SEMAINE), 'Arrêt Machine (minutes)'].sum()
    
    # Mettre à jour les lignes de totaux
    df.at[6, 'Temps D\'ouverture'] = total_minutes
    df.at[7, 'Temps D\'ouverture'] = total_minutes / 60
    df.at[6, 'Arrêt Machine (minutes)'] = total_arret
    df.at[7, 'Arrêt Machine (minutes)'] = total_arret / 60
    
    # Calculer la disponibilité (en %)
    for i in range(6):  # Pour chaque jour
        temps_ouvert = df.at[i, 'Temps D\'ouverture']
        arret = df.at[i, 'Arrêt Machine (minutes)']
        if temps_ouvert > 0:
            dispo = ((temps_ouvert - arret) / temps_ouvert) * 100
            df.at[i, 'Disponibilité Machine'] = f"{dispo:.1f}%"
    
    # Sauvegarder dans le fichier Excel
    df.to_excel(file_path, index=False, sheet_name=machine_name)

try:
    # Initialisation des fichiers pour chaque machine
    init_machine_file(G11_FILE, "G11")
    init_machine_file(G20_FILE, "G20")
    init_machine_file(C12_FILE, "C12")
    
    # Sélection de la machine
    machine = st.selectbox("Sélectionnez la machine", ["G11", "G20", "C12"])
    
    # Charger les données de la machine sélectionnée
    if machine == "G11":
        df = load_machine_data(G11_FILE)
        file_path = G11_FILE
    elif machine == "G20":
        df = load_machine_data(G20_FILE)
        file_path = G20_FILE
    else:  # C12
        df = load_machine_data(C12_FILE)
        file_path = C12_FILE
    
    # Afficher le formulaire
    st.header(f"Suivi de la machine {machine}")
    
    # Afficher les données existantes dans un tableau éditable
    st.subheader("Données de la semaine")
    
    # Créer un formulaire pour l'édition
    with st.form(f"form_{machine}"):
        # Créer une copie du dataframe pour l'édition
        edited_df = df.copy()
        
        # Afficher les données dans un éditeur de données avec une configuration simplifiée
        st.write("Modifiez les données ci-dessous :")
        
        # Pour chaque jour de la semaine
        for i, jour in enumerate(JOURS_SEMAINE):
            st.subheader(jour)
            
            # Créer des colonnes pour une meilleure mise en page
            col1, col2, col3 = st.columns([1, 1, 2])
            
            with col1:
                # Temps d'ouverture
                edited_df.at[i, 'Temps D\'ouverture'] = st.number_input(
                    "Temps d'ouverture (min)",
                    min_value=0,
                    max_value=1440,
                    value=int(edited_df.at[i, 'Temps D\'ouverture']),
                    step=1,
                    key=f"ouverture_{machine}_{i}"
                )
            
            with col2:
                # Arrêt machine
                edited_df.at[i, 'Arrêt Machine (minutes)'] = st.number_input(
                    "Arrêt (min)",
                    min_value=0,
                    max_value=1440,
                    value=int(edited_df.at[i, 'Arrêt Machine (minutes)']),
                    step=1,
                    key=f"arret_{machine}_{i}"
                )
                
                # Calculer et afficher la disponibilité
                temps_ouvert = edited_df.at[i, 'Temps D\'ouverture']
                arret = edited_df.at[i, 'Arrêt Machine (minutes)']
                dispo = ((temps_ouvert - arret) / temps_ouvert * 100) if temps_ouvert > 0 else 0
                st.metric("Disponibilité", f"{dispo:.1f}%")
            
            with col3:
                # Interventions
                edited_df.at[i, 'Interventions'] = st.text_area(
                    "Interventions",
                    value=str(edited_df.at[i, 'Interventions'] or ''),
                    key=f"interv_{machine}_{i}",
                    height=100
                )
            
            st.markdown("---")
        
        # Afficher les totaux en lecture seule
        st.subheader("Totaux")
        total_minutes = edited_df.loc[edited_df['Jour'].isin(JOURS_SEMAINE), 'Temps D\'ouverture'].sum()
        total_arret = edited_df.loc[edited_df['Jour'].isin(JOURS_SEMAINE), 'Arrêt Machine (minutes)'].sum()
        dispo = ((total_minutes - total_arret) / total_minutes * 100) if total_minutes > 0 else 0
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Temps d'ouverture total", f"{total_minutes} min")
        with col2:
            st.metric("Temps d'arrêt total", f"{total_arret} min")
        with col3:
            st.metric("Disponibilité globale", f"{dispo:.1f}%")
        
        # Bouton de soumission
        submitted = st.form_submit_button("Enregistrer les modifications")
        
        if submitted:
            try:
                # Sauvegarder les modifications
                save_machine_data(file_path, edited_df, machine)
                st.success("Données enregistrées avec succès!")
                
                # Recharger les données mises à jour
                if machine == "G11":
                    df = load_machine_data(G11_FILE)
                elif machine == "G20":
                    df = load_machine_data(G20_FILE)
                else:  # C12
                    df = load_machine_data(C12_FILE)
                
            except Exception as e:
                st.error(f"Erreur lors de l'enregistrement : {e}")
    
    # Afficher un résumé dans la barre latérale
    st.sidebar.title(f"Résumé - {machine}")
    if df is not None:
        # Calculer le temps d'arrêt total
        total_arret = df.loc[df['Jour'].isin(JOURS_SEMAINE), 'Arrêt Machine (minutes)'].sum()
        total_heures = total_arret / 60
        dispo = ((8640 - total_arret) / 8640) * 100
        
        st.sidebar.metric("Temps d'arrêt total", f"{total_arret} min")
        st.sidebar.metric("Temps d'arrêt (heures)", f"{total_heures:.1f} h")
        st.sidebar.metric("Disponibilité moyenne", f"{dispo:.1f}%")
        
        # Afficher les interventions
        interventions = df[df['Interventions'].notna() & (df['Interventions'] != '')]
        if not interventions.empty:
            st.sidebar.subheader("Interventions enregistrées")
            for _, row in interventions.iterrows():
                if row['Jour'] in JOURS_SEMAINE:  # Ne pas afficher les totaux
                    st.sidebar.info(f"**{row['Jour']}**: {row['Interventions']}")

except Exception as e:
    st.error(f"Une erreur est survenue : {e}")
    st.stop()
