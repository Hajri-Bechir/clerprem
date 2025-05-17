
import streamlit as st
import pandas as pd
import os
from openpyxl import load_workbook

# This page is for sheets: CT1 to CT4
SHEET_NAME = "CT1_CT4"
# Use a dedicated Excel file for CT1_CT4
EXCEL_FILE_PATH = r"C:/Users/White/Desktop/clerprem-project/data/CT1_CT4_data.xlsx"  # Dedicated file for CT1_CT4

st.set_page_config(page_title=f"Input: Ct3", page_icon="📝")
st.title(f"📝 Data Input for Sheet: {SHEET_NAME}")

# Create the Excel file with headers if it doesn't exist
if not os.path.exists(EXCEL_FILE_PATH):
    df = pd.DataFrame(columns=['Project', 'Familles', 'Rump up journalier', 'Objecti Semaine', 
                             'Réaliser', '%', 'Commentaires', 'Lundi', 'Mardi', 'Mercredi', 
                             'Jeudi', 'Vendredi', 'Samedi'])
    df.to_excel(EXCEL_FILE_PATH, index=False)
    st.success(f"Created new Excel file at: {EXCEL_FILE_PATH}")
elif not os.access(EXCEL_FILE_PATH, os.W_OK):
    st.error(f"Cannot write to file: {EXCEL_FILE_PATH}. Please check file permissions.")
    st.stop()

try:
    # Define the exact column order as in the Excel sheet
    columns = ['Project', 'Familles', 'Rump up journalier', 'Objecti Semaine', 'Réaliser', '%', 'Commentaires',
              'Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi']

    with st.form(key=f"data_input_form_{SHEET_NAME.replace(' ', '_')}"):
        # Project and Family Selection
        st.subheader("1. Project Selection")
        col1, col2 = st.columns(2)
        with col1:
            # Define dropdown options for Project and Familles
            project_familles = {
                'AB3/Q2': 'AB3/Q2 Agrafage',
                'C8': [
                    'Ablage C8 Hinten',
                    'Kopfkasten C8 Hinten',
                    'Polsterriegel C8 Hinten',
                    'Deckel C8 Hinten',
                    'Deckel C8 Vorne'
                ],
                'C-BEV': 'C-BEV',
                'D5': 'D5 Low',
                'Seipo': 'Seipo E6',
                'SE38': 'Mal Vorne Seat',
                'EQ 5': 'SEIPO EQ 5',
                'Seipo SE': 'SEIPO SE38',
                'SEIPO SK': [
                    'SEIPO SK38',
                    'SEIPO SK316'
                ],
                'SEIPO VW': [
                    'SEIPO VW380',
                    'SEIPO VW 382'
                ],
                'T-ROC': 'T-ROC',
                'Renault': 'HHN'
            }
            selected_project = st.selectbox("Project", list(project_familles.keys()), 
                                         help="Select the project for this entry")
            
        with col2:
            if isinstance(project_familles[selected_project], list):
                selected_familles = st.selectbox("Familles", project_familles[selected_project],
                                                help="Select the specific family within the project")
            else:
                selected_familles = project_familles[selected_project]
            
        new_row_data = {col: None for col in columns}
        new_row_data['Project'] = selected_project
        new_row_data['Familles'] = selected_familles
        
        # Weekly Metrics Section
        st.subheader("2. Weekly Metrics")
        st.markdown("<div style='padding: 1em; border-radius:10px; border:2px solid #4F8BF9; background-color:#F6FAFF; margin-bottom:1em;'>", unsafe_allow_html=True)
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            new_row_data['Rump up journalier'] = st.number_input("🔵 Rump up journalier", step=1, format="%d",
                                                               help="Daily target quantity")
        with col2:
            new_row_data['Objecti Semaine'] = st.number_input("🟢 Objecti Semaine", step=1, format="%d",
                                                            help="Weekly target quantity")
        with col3:
            new_row_data['Réaliser'] = st.number_input("🟣 Réaliser", step=1, format="%d",
                                                     help="Actual quantity achieved")
        with col4:
            if new_row_data['Objecti Semaine'] != 0:
                calc_percent = (new_row_data['Réaliser'] / new_row_data['Objecti Semaine']) * 100
            else:
                calc_percent = 0.0
            st.markdown(f"<b>🎯 %: <span style='color:#2E8B57'>{calc_percent:.2f}%</span></b>", unsafe_allow_html=True)
            new_row_data['%'] = calc_percent
        st.markdown("</div>", unsafe_allow_html=True)
        
        # Comments Section
        st.subheader("3. Comments")
        new_row_data['Commentaires'] = st.text_area("Commentaires",
                                                  help="Add any relevant notes or comments about this entry")
        
        # Daily Production Section
        st.subheader("4. Daily Production")
        days = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi"]
        cols = st.columns(len(days))
        for i, day in enumerate(days):
            with cols[i]:
                new_row_data[day] = st.number_input(day, step=1, format="%d", 
                                                  key=f"daily_{day}",
                                                  help=f"Quantity produced on {day}")
        
        # Submit Button
        submit_button = st.form_submit_button(label=f"Add Row to Sheet: {SHEET_NAME}")

    if submit_button:
        try:
            # Read existing data
            df_existing = pd.read_excel(EXCEL_FILE_PATH)
            
            # Create a new DataFrame with the new row
            new_row_df = pd.DataFrame([new_row_data])
            
            # Append the new row to the existing data
            df_updated = pd.concat([df_existing, new_row_df], ignore_index=True)
            
            # Save back to Excel
            df_updated.to_excel(EXCEL_FILE_PATH, index=False)
            st.success("Données enregistrées avec succès!")

        except Exception as e:
            st.error(f"Erreur lors de l'enregistrement des données : {e}")

    # Show existing data in sidebar
    st.sidebar.markdown("--- ")
    if st.sidebar.checkbox(f"Afficher les données existantes pour {SHEET_NAME}?", key=f"show_data_{SHEET_NAME.replace(' ', '_')}"):
        try:
            df_display = pd.read_excel(EXCEL_FILE_PATH)
            st.sidebar.subheader(f"Données actuelles dans '{SHEET_NAME}' (5 premières lignes)")
            st.sidebar.dataframe(df_display.head())
        except Exception as e:
            st.sidebar.error(f"Impossible d'afficher les données : {e}")

except FileNotFoundError:
    st.error(f"ERREUR: Fichier Excel introuvable à l'emplacement {EXCEL_FILE_PATH}.")
except Exception as e:
    st.error(f"Une erreur inattendue s'est produite : {e}")
    st.info("Veuillez vérifier que le fichier Excel n'est pas ouvert dans un autre programme.")

