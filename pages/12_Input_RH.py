import streamlit as st
import pandas as pd
import os
from datetime import datetime

# Configuration de la page
st.set_page_config(page_title="Suivi RH - Effectifs", page_icon="👥")
st.title("👥 Suivi des Ressources Humaines")

# Chemins des fichiers Excel
RH_FILE = r"C:/Users/White/Desktop/clerprem-project/data/Suivi_RH.xlsx"

# Jours de la semaine
JOURS_SEMAINE = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi"]
SECTIONS = ["CT2-CT8", "CT3", "CT5", "CT6", "CT1", "CT4", "CT7", "CT9"]

# Mapping pour l'affichage des sections (pour l'interface utilisateur)
SECTION_DISPLAY = {
    "CT2-CT8": "CT2/CT8",
    "CT3": "CT3",
    "CT5": "CT5",
    "CT6": "CT6",
    "CT1": "CT1",
    "CT4": "CT4",
    "CT7": "CT7",
    "CT9": "CT9"
}

# Structure des données par section
def init_section_data(section):
    data = {
        'Jour': JOURS_SEMAINE + ['Total Semaine'],
        'Operateurs Present': [0] * 7,
        'Operateurs Absent': [0] * 7,
        '% Absent': ['0%'] * 7,
        'Opérateurs Sortie': [0] * 7,
        '% Sortie': ['0%'] * 7,
        'Opérateurs Embauchés': [0] * 7,
        '% Embauchés': ['0%'] * 7,
        'Section': [section] * 7
    }
    return pd.DataFrame(data)

# Fonction pour charger les données RH
def load_rh_data():
    if not os.path.exists(RH_FILE):
        # Créer un nouveau fichier avec toutes les sections
        with pd.ExcelWriter(RH_FILE, engine='openpyxl') as writer:
            for section in SECTIONS:
                df_section = init_section_data(section)
                df_section.to_excel(writer, sheet_name=section, index=False)
    
    # Vérifier et ajouter les sections manquantes
    with pd.ExcelFile(RH_FILE) as xls:
        existing_sheets = xls.sheet_names
    
    # Créer les feuilles manquantes si nécessaire
    with pd.ExcelWriter(RH_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        for section in SECTIONS:
            if section not in existing_sheets:
                df_section = init_section_data(section)
                df_section.to_excel(writer, sheet_name=section, index=False)
    
    return None

# Fonction pour sauvegarder les données RH
def save_rh_data(df, section):
    # Mettre à jour les totaux
    for col in ['Operateurs Present', 'Operateurs Absent', 'Opérateurs Sortie', 'Opérateurs Embauchés']:
        df.at[6, col] = df[col].iloc[:6].sum()
    
    # Calculer les pourcentages
    for i in range(7):  # Pour chaque jour + total
        total = df.at[i, 'Operateurs Present'] + df.at[i, 'Operateurs Absent']
        
        # Éviter la division par zéro
        if total > 0:
            # Pourcentage d'absence
            df.at[i, '% Absent'] = f"{df.at[i, 'Operateurs Absent'] / total * 100:.1f}%"
            
            # Pourcentage de sortie
            df.at[i, '% Sortie'] = f"{df.at[i, 'Opérateurs Sortie'] / total * 100:.1f}%"
            
            # Pourcentage d'embauche
            df.at[i, '% Embauchés'] = f"{df.at[i, 'Opérateurs Embauchés'] / total * 100:.1f}%"
        else:
            df.at[i, '% Absent'] = "0%"
            df.at[i, '% Sortie'] = "0%"
            df.at[i, '% Embauchés'] = "0%"
    
    # Sauvegarder dans le fichier Excel
    with pd.ExcelWriter(RH_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name=section, index=False)

try:
    # Charger ou initialiser les données
    load_rh_data()
    
    # Sélection de la section
    section_display = st.sidebar.selectbox("Sélectionnez la section", list(SECTION_DISPLAY.values()))
    # Récupérer la clé de section correspondante
    section = next(key for key, value in SECTION_DISPLAY.items() if value == section_display)
    
    # Charger les données de la section
    try:
        df_section = pd.read_excel(RH_FILE, sheet_name=section)
    except Exception as e:
        # Si la feuille n'existe pas, la créer
        df_section = init_section_data(section)
        with pd.ExcelWriter(RH_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_section.to_excel(writer, sheet_name=section, index=False)
    
    st.header(f"Section {SECTION_DISPLAY[section]}")
    
    # Créer un formulaire pour l'édition
    with st.form(f"form_rh_{section}"):
        st.subheader("Saisie des données")
        
        # Pour chaque jour de la semaine
        for i, jour in enumerate(JOURS_SEMAINE):
            st.markdown(f"### {jour}")
            
            # Créer des colonnes pour une meilleure mise en page
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                # Opérateurs présents
                df_section.at[i, 'Operateurs Present'] = st.number_input(
                    "Opérateurs présents",
                    min_value=0,
                    value=int(df_section.at[i, 'Operateurs Present']),
                    step=1,
                    key=f"present_{section}_{i}"
                )
            
            with col2:
                # Opérateurs absents
                df_section.at[i, 'Operateurs Absent'] = st.number_input(
                    "Opérateurs absents",
                    min_value=0,
                    value=int(df_section.at[i, 'Operateurs Absent']),
                    step=1,
                    key=f"absent_{section}_{i}"
                )
            
            with col3:
                # Opérateurs sortis
                df_section.at[i, 'Opérateurs Sortie'] = st.number_input(
                    "Sorties",
                    min_value=0,
                    value=int(df_section.at[i, 'Opérateurs Sortie']),
                    step=1,
                    key=f"sortie_{section}_{i}"
                )
            
            with col4:
                # Opérateurs embauchés
                df_section.at[i, 'Opérateurs Embauchés'] = st.number_input(
                    "Embauches",
                    min_value=0,
                    value=int(df_section.at[i, 'Opérateurs Embauchés']),
                    step=1,
                    key=f"embauche_{section}_{i}"
                )
            
            st.markdown("---")
        
        # Bouton de soumission
        submitted = st.form_submit_button("Enregistrer les modifications")
        
        if submitted:
            try:
                save_rh_data(df_section, section)
                st.success("Données enregistrées avec succès!")
                
                # Recharger les données mises à jour
                df_section = pd.read_excel(RH_FILE, sheet_name=section)
                
            except Exception as e:
                st.error(f"Erreur lors de l'enregistrement : {e}")
    
    # Afficher les totaux dans la barre latérale
    st.sidebar.title("Résumé")
    
    # Calculer les totaux
    total_presents = df_section.loc[df_section['Jour'].isin(JOURS_SEMAINE), 'Operateurs Present'].sum()
    total_absents = df_section.loc[df_section['Jour'].isin(JOURS_SEMAINE), 'Operateurs Absent'].sum()
    total_sorties = df_section.loc[df_section['Jour'].isin(JOURS_SEMAINE), 'Opérateurs Sortie'].sum()
    total_embauches = df_section.loc[df_section['Jour'].isin(JOURS_SEMAINE), 'Opérateurs Embauchés'].sum()
    
    # Afficher les indicateurs clés
    st.sidebar.metric("Effectif moyen présent", f"{total_presents/6:.1f}")
    st.sidebar.metric("Taux d'absence", f"{total_absents/(total_presents + total_absents)*100 if (total_presents + total_absents) > 0 else 0:.1f}%")
    st.sidebar.metric("Taux de rotation", f"{total_sorties/(total_presents + total_absents)*100 if (total_presents + total_absents) > 0 else 0:.1f}%")
    st.sidebar.metric("Taux d'embauche", f"{total_embauches/(total_presents + total_absents)*100 if (total_presents + total_absents) > 0 else 0:.1f}%")
    
    # Afficher les données sous forme de tableau
    st.subheader("Synthèse des données")
    st.dataframe(df_section, use_container_width=True, hide_index=True)

except Exception as e:
    st.error(f"Une erreur est survenue : {e}")
    st.stop()
