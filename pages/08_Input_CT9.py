
import streamlit as st
import pandas as pd
import os
from openpyxl import load_workbook

# This page is for sheet: CT9
SHEET_NAME = "CT9"
# Use a dedicated Excel file for CT9
EXCEL_FILE_PATH = r"C:/Users/White/Desktop/clerprem-project/data/CT9_data.xlsx"  # Dedicated file for CT9

st.set_page_config(page_title=f"Input: Ct9", page_icon="📝")
st.title(f"📝 Data Input for Sheet: {SHEET_NAME}")

# Create the Excel file with headers if it doesn't exist
if not os.path.exists(EXCEL_FILE_PATH):
    df = pd.DataFrame(columns=['Project', 'Familles', 'Rump up journalier', 'Objecti Semaine', 
                             'Réaliser', '%', 'Commentaires', 'Lundi', 'Mardi', 'Mercredi', 
                             'Jeudi', 'Vendredi', 'Samedi'])
    df.to_excel(EXCEL_FILE_PATH, index=False, sheet_name=SHEET_NAME)
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
                'BMW': 'G6X',
                'EQ 5': 'SEIPO EQ 5'
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

        # Add other fields
        st.subheader("2. Enter Data")
        col1, col2, col3 = st.columns(3)
        with col1:
            new_row_data['Rump up journalier'] = st.number_input("Rump up journalier", value=0, min_value=0, step=1)
            new_row_data['Objecti Semaine'] = st.number_input("Objectif Semaine", value=0, min_value=0, step=1)
            new_row_data['Réaliser'] = st.number_input("Réaliser", value=0, min_value=0, step=1)
            new_row_data['%'] = st.number_input("%", value=0.0, min_value=0.0, max_value=100.0, step=0.1)
        
        new_row_data['Commentaires'] = st.text_area("Commentaires")
        
        st.subheader("3. Jours de la semaine")
        days = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi']
        cols = st.columns(len(days))
        for i, day in enumerate(days):
            with cols[i]:
                new_row_data[day] = st.number_input(day, value=0, min_value=0, step=1, key=f"{day}_{SHEET_NAME}")
        
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

            st.success(f"Data successfully added to sheet '{SHEET_NAME}' and saved!")
            

        except Exception as e:
            st.error(f"Error saving data to Excel for sheet '{SHEET_NAME}': {e}")

    st.sidebar.markdown("--- ")
    if st.sidebar.checkbox(f"Show existing data for {SHEET_NAME}?", key=f"show_data_{SHEET_NAME.replace(' ', '_')}"):
        try:
            df_display = pd.read_excel(EXCEL_FILE_PATH)
            st.sidebar.subheader(f"Current Data in '{SHEET_NAME}' (first 5 rows)")
            st.sidebar.dataframe(df_display.head())
        except Exception as e:
            st.sidebar.error(f"Could not display data for '{SHEET_NAME}': {e}")

except FileNotFoundError:
    st.error(f"ERREUR: Fichier Excel introuvable à l'emplacement {EXCEL_FILE_PATH}. Veuillez vérifier le chemin d'accès.")
except KeyError as e:
    st.error(f"ERREUR: Onglet '{SHEET_NAME}' introuvable dans le fichier Excel. {e}")
    st.error(f"Erreur détaillée : {str(e)}")
    import traceback
    st.text(traceback.format_exc())
except Exception as e:
    st.error(f"An unexpected error occurred while processing sheet '{SHEET_NAME}': {e}")
    st.info("Please ensure the Excel file and sheet exist and are correctly formatted.")
