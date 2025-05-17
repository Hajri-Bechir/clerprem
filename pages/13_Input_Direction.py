import streamlit as st
import pandas as pd
import os
from openpyxl import load_workbook

# This page is for sheet: Direction
SHEET_NAME = "Direction"
# Use a dedicated Excel file for Direction data, similar to CT9
EXCEL_FILE_PATH_IN_PAGE = r"C:/Users/White/Desktop/clerprem-project/data/Direction_data.xlsx" # Dedicated file for Direction

st.set_page_config(page_title=f"Input: Direction", page_icon="📊")
st.title(f"📊 Direction Data Input")

# Create the Excel file with headers if it doesn't exist
if not os.path.exists(EXCEL_FILE_PATH_IN_PAGE):
    # Create the directory if it doesn't exist
    os.makedirs(os.path.dirname(EXCEL_FILE_PATH_IN_PAGE), exist_ok=True)
    
    # Create a new Excel file with the required columns
    columns = ['Project', 'Familles', 'Rump up journalier', 'Objecti Semaine', 'Réaliser', '%',
              'Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi',
              'Production', 'Operateurs Present', 'Operateurs Absent', '% Absent',
              'Opérateurs Sortie', '% Sortie', 'Opérateurs Embauchés', '% Embauchés',
              'Maintenance_Lundi_Temps_Ouverture', 'Maintenance_Lundi_Arret_Machine', 'Maintenance_Lundi_Disponibilite',
              'Maintenance_Mardi_Temps_Ouverture', 'Maintenance_Mardi_Arret_Machine', 'Maintenance_Mardi_Disponibilite',
              'Maintenance_Mercredi_Temps_Ouverture', 'Maintenance_Mercredi_Arret_Machine', 'Maintenance_Mercredi_Disponibilite',
              'Maintenance_Jeudi_Temps_Ouverture', 'Maintenance_Jeudi_Arret_Machine', 'Maintenance_Jeudi_Disponibilite',
              'Maintenance_Vendredi_Temps_Ouverture', 'Maintenance_Vendredi_Arret_Machine', 'Maintenance_Vendredi_Disponibilite',
              'Maintenance_Samedi_Temps_Ouverture', 'Maintenance_Samedi_Arret_Machine', 'Maintenance_Samedi_Disponibilite',
              'Maintenance_Total_Minutes_Ouverture', 'Maintenance_Total_Minutes_Arret', 'Maintenance_Total_Disponibilite',
              'Maintenance_Total_Heures_Ouverture', 'Maintenance_Total_Heures_Arret', 'Maintenance_Total_Disponibilite_Pct',
              'Visite_Date', 'Visite_Semaine', 'Visite_Motif', 'Visite_Qui']
    
    df = pd.DataFrame(columns=columns)
    df.to_excel(EXCEL_FILE_PATH_IN_PAGE, index=False)
    st.success(f"Created new Excel file at: {EXCEL_FILE_PATH_IN_PAGE}")
elif not os.access(EXCEL_FILE_PATH_IN_PAGE, os.W_OK):
    st.error(f"Cannot write to file: {EXCEL_FILE_PATH_IN_PAGE}. Please check file permissions.")
    st.stop()

try:
    # Define project and families mapping
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
        'Crafter': 'Crafter',
        'D5': 'D5 Low',
        'Seipo': 'Seipo E6',
        'BMW': 'G6X',
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
        'T-ROC': 'T-ROC'
    }
    
    with st.form(key=f"data_input_form_{SHEET_NAME.replace(' ', '_')}"):
        # Project and Family Selection
        st.subheader("1. Project and Family Selection")
        col1, col2 = st.columns(2)
        with col1:
            selected_project = st.selectbox("Project", list(project_familles.keys()), 
                                         help="Select the project for this entry")
            
        with col2:
            if isinstance(project_familles[selected_project], list):
                selected_familles = st.selectbox("Familles", project_familles[selected_project],
                                               help="Select the specific family within the project")
            else:
                selected_familles = project_familles[selected_project]
        
        # Initialize new row data dictionary
        new_row_data = {}
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
            if new_row_data.get('Objecti Semaine', 0) != 0:
                calc_percent = (new_row_data.get('Réaliser', 0) / new_row_data.get('Objecti Semaine', 1)) * 100
            else:
                calc_percent = 0.0
            st.markdown(f"<b>🎯 %: <span style='color:#2E8B57'>{calc_percent:.2f}%</span></b>", unsafe_allow_html=True)
            new_row_data['%'] = calc_percent
        st.markdown("</div>", unsafe_allow_html=True)
        
        # Daily Production Section
        st.subheader("3. Daily Production")
        days = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi"]
        cols = st.columns(len(days))
        for i, day in enumerate(days):
            with cols[i]:
                new_row_data[day] = st.number_input(day, step=1, format="%d", 
                                                 key=f"daily_{day}",
                                                 help=f"Quantity produced on {day}")
        
        # RH Data Section
        st.subheader("4. HR Data")
        new_row_data['Production'] = st.number_input("Production", step=1, format="%d", 
                                                 help="Total production")
        
        # Create two rows of HR metrics
        st.markdown("<div style='padding: 1em; border-radius:10px; border:2px solid #4F8BF9; background-color:#F6FAFF; margin-bottom:1em;'>", unsafe_allow_html=True)
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            new_row_data['Operateurs Present'] = st.number_input("👨‍💼 Operateurs Present", step=1, format="%d",
                                                              help="Number of operators present")
        with col2:
            new_row_data['Operateurs Absent'] = st.number_input("🚶 Operateurs Absent", step=1, format="%d",
                                                             help="Number of operators absent")
        with col3:
            if (new_row_data.get('Operateurs Present', 0) + new_row_data.get('Operateurs Absent', 0)) > 0:
                absent_percent = (new_row_data.get('Operateurs Absent', 0) / 
                                 (new_row_data.get('Operateurs Present', 0) + new_row_data.get('Operateurs Absent', 0))) * 100
            else:
                absent_percent = 0.0
            st.markdown(f"<b>% Absent: <span style='color:#E74C3C'>{absent_percent:.2f}%</span></b>", unsafe_allow_html=True)
            new_row_data['% Absent'] = absent_percent
        
        st.markdown("</div>", unsafe_allow_html=True)
        
        st.markdown("<div style='padding: 1em; border-radius:10px; border:2px solid #4F8BF9; background-color:#F6FAFF; margin-bottom:1em;'>", unsafe_allow_html=True)
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            new_row_data['Opérateurs Sortie'] = st.number_input("🚪 Opérateurs Sortie", step=1, format="%d",
                                                             help="Number of operators who left")
        with col2:
            if (new_row_data.get('Operateurs Present', 0) + new_row_data.get('Operateurs Absent', 0)) > 0:
                sortie_percent = (new_row_data.get('Opérateurs Sortie', 0) / 
                                 (new_row_data.get('Operateurs Present', 0) + new_row_data.get('Operateurs Absent', 0))) * 100
            else:
                sortie_percent = 0.0
            st.markdown(f"<b>% Sortie: <span style='color:#E74C3C'>{sortie_percent:.2f}%</span></b>", unsafe_allow_html=True)
            new_row_data['% Sortie'] = sortie_percent
        
        with col3:
            new_row_data['Opérateurs Embauchés'] = st.number_input("👋 Opérateurs Embauchés", step=1, format="%d",
                                                                help="Number of operators hired")
        with col4:
            if (new_row_data.get('Operateurs Present', 0) + new_row_data.get('Operateurs Absent', 0)) > 0:
                embauches_percent = (new_row_data.get('Opérateurs Embauchés', 0) / 
                                    (new_row_data.get('Operateurs Present', 0) + new_row_data.get('Operateurs Absent', 0))) * 100
            else:
                embauches_percent = 0.0
            st.markdown(f"<b>% Embauchés: <span style='color:#2E8B57'>{embauches_percent:.2f}%</span></b>", unsafe_allow_html=True)
            new_row_data['% Embauchés'] = embauches_percent
        
        st.markdown("</div>", unsafe_allow_html=True)
        
        # Maintenance Section
        st.subheader("5. Maintenance - Machine Injection")
        
        # Create a DataFrame to hold maintenance data temporarily for display and calculations
        maintenance_data = pd.DataFrame(
            index=["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Total Minutes", "Total Heures"],
            columns=["Temps D'ouverture", "Arrêt Machine", "Disponibilité Machine"]
        )
        
        # Fill with zeros/empty
        maintenance_data.fillna(0, inplace=True)
        maintenance_data.loc["Total Heures", "Disponibilité Machine"] = "#DIV/0!"
        
        # Create a container with styling
        st.markdown("<div style='padding: 1em; border-radius:10px; border:2px solid #4F8BF9; background-color:#F6FAFF; margin-bottom:1em;'>", unsafe_allow_html=True)
        
        # Create a table-like structure for maintenance data
        days = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi"]
        
        # Column headers
        col1, col2, col3, col4 = st.columns([1, 1.5, 1.5, 1.5])
        with col1:
            st.markdown("<b>Jour</b>", unsafe_allow_html=True)
        with col2:
            st.markdown("<b>Temps D'ouverture (min)</b>", unsafe_allow_html=True)
        with col3:
            st.markdown("<b>Arrêt Machine (min)</b>", unsafe_allow_html=True)
        with col4:
            st.markdown("<b>Disponibilité Machine</b>", unsafe_allow_html=True)
        
        # Input rows for each day
        total_temps_ouverture = 0
        total_arret_machine = 0
        
        for day in days:
            col1, col2, col3, col4 = st.columns([1, 1.5, 1.5, 1.5])
            
            with col1:
                st.markdown(f"<b>{day}</b>", unsafe_allow_html=True)
            
            with col2:
                temps_key = f"maintenance_{day.lower()}_temps_ouverture"
                temps_ouverture = st.number_input(
                    "", 
                    min_value=0, 
                    step=1, 
                    key=temps_key,
                    label_visibility="collapsed"
                )
                maintenance_data.loc[day, "Temps D'ouverture"] = temps_ouverture
                new_row_data[f"Maintenance_{day}_Temps_Ouverture"] = temps_ouverture
                total_temps_ouverture += temps_ouverture
            
            with col3:
                arret_key = f"maintenance_{day.lower()}_arret_machine"
                arret_machine = st.number_input(
                    "", 
                    min_value=0, 
                    step=1, 
                    key=arret_key,
                    label_visibility="collapsed"
                )
                maintenance_data.loc[day, "Arrêt Machine"] = arret_machine
                new_row_data[f"Maintenance_{day}_Arret_Machine"] = arret_machine
                total_arret_machine += arret_machine
            
            with col4:
                if temps_ouverture > 0:
                    disponibilite = ((temps_ouverture - arret_machine) / temps_ouverture) * 100
                    disponibilite_str = f"{disponibilite:.2f}%"
                else:
                    disponibilite = 0
                    disponibilite_str = "0.00%"
                
                st.markdown(f"<span>{disponibilite_str}</span>", unsafe_allow_html=True)
                maintenance_data.loc[day, "Disponibilité Machine"] = disponibilite_str
                new_row_data[f"Maintenance_{day}_Disponibilite"] = disponibilite
        
        # Display totals
        st.markdown("<hr style='margin: 0.5em 0;'>", unsafe_allow_html=True)
        
        # Total Minutes row
        col1, col2, col3, col4 = st.columns([1, 1.5, 1.5, 1.5])
        with col1:
            st.markdown("<b>Total Minutes</b>", unsafe_allow_html=True)
        with col2:
            st.markdown(f"<b>{total_temps_ouverture}</b>", unsafe_allow_html=True)
            maintenance_data.loc["Total Minutes", "Temps D'ouverture"] = total_temps_ouverture
            new_row_data["Maintenance_Total_Minutes_Ouverture"] = total_temps_ouverture
        with col3:
            st.markdown(f"<b>{total_arret_machine}</b>", unsafe_allow_html=True)
            maintenance_data.loc["Total Minutes", "Arrêt Machine"] = total_arret_machine
            new_row_data["Maintenance_Total_Minutes_Arret"] = total_arret_machine
        with col4:
            if total_temps_ouverture > 0:
                total_disponibilite = ((total_temps_ouverture - total_arret_machine) / total_temps_ouverture) * 100
                total_disponibilite_str = f"{total_disponibilite:.2f}%"
            else:
                total_disponibilite = 0
                total_disponibilite_str = "0.00%"
            
            st.markdown(f"<b>{total_disponibilite_str}</b>", unsafe_allow_html=True)
            maintenance_data.loc["Total Minutes", "Disponibilité Machine"] = total_disponibilite_str
            new_row_data["Maintenance_Total_Disponibilite"] = total_disponibilite
        
        # Total Hours row
        col1, col2, col3, col4 = st.columns([1, 1.5, 1.5, 1.5])
        with col1:
            st.markdown("<b>Total Heures</b>", unsafe_allow_html=True)
        with col2:
            total_heures_ouverture = total_temps_ouverture / 60 if total_temps_ouverture > 0 else 0
            st.markdown(f"<b>{total_heures_ouverture:.2f}</b>", unsafe_allow_html=True)
            maintenance_data.loc["Total Heures", "Temps D'ouverture"] = total_heures_ouverture
            new_row_data["Maintenance_Total_Heures_Ouverture"] = total_heures_ouverture
        with col3:
            total_heures_arret = total_arret_machine / 60 if total_arret_machine > 0 else 0
            st.markdown(f"<b>{total_heures_arret:.2f}</b>", unsafe_allow_html=True)
            maintenance_data.loc["Total Heures", "Arrêt Machine"] = total_heures_arret
            new_row_data["Maintenance_Total_Heures_Arret"] = total_heures_arret
        with col4:
            st.markdown(f"<b>{total_disponibilite_str}</b>", unsafe_allow_html=True)
            maintenance_data.loc["Total Heures", "Disponibilité Machine"] = total_disponibilite_str
            new_row_data["Maintenance_Total_Disponibilite_Pct"] = total_disponibilite
        
        st.markdown("</div>", unsafe_allow_html=True)
        
        # Visite Client ou Autre Section
        st.subheader("6. Visite Client ou Autre")
        st.markdown("<div style='padding: 1em; border-radius:10px; border:2px solid #4F8BF9; background-color:#F6FAFF; margin-bottom:1em;'>", unsafe_allow_html=True)
        
        # Table headers
        col1, col2, col3, col4 = st.columns([1, 1, 2, 1])
        with col1:
            st.markdown("<b>Date</b>", unsafe_allow_html=True)
        with col2:
            st.markdown("<b>Semaine</b>", unsafe_allow_html=True)
        with col3:
            st.markdown("<b>Motif Visite</b>", unsafe_allow_html=True)
        with col4:
            st.markdown("<b>Qui ?</b>", unsafe_allow_html=True)
        
        # Input row
        col1, col2, col3, col4 = st.columns([1, 1, 2, 1])
        
        with col1:
            visite_date = st.date_input(
                "",
                value=None,
                key="visite_date",
                label_visibility="collapsed"
            )
            new_row_data['Visite_Date'] = visite_date
        
        with col2:
            # Calculate week number from date if date is selected
            if visite_date:
                week_num = visite_date.isocalendar()[1]
                visite_semaine = st.text_input(
                    "", 
                    value=f"S{week_num}",
                    key="visite_semaine",
                    label_visibility="collapsed"
                )
            else:
                visite_semaine = st.text_input(
                    "", 
                    placeholder="ex: S20",
                    key="visite_semaine",
                    label_visibility="collapsed"
                )
            new_row_data['Visite_Semaine'] = visite_semaine
        
        with col3:
            visite_motif = st.text_area(
                "", 
                placeholder="Motif de la visite",
                key="visite_motif",
                label_visibility="collapsed",
                height=100
            )
            new_row_data['Visite_Motif'] = visite_motif
        
        with col4:
            visite_qui = st.text_input(
                "", 
                placeholder="Nom(s)",
                key="visite_qui",
                label_visibility="collapsed"
            )
            new_row_data['Visite_Qui'] = visite_qui
        
        st.markdown("</div>", unsafe_allow_html=True)
        
        # Submit Button
        submit_button = st.form_submit_button(label=f"Add Row to Sheet: {SHEET_NAME}")
        
    if submit_button:
        try:
            # Read existing data
            try:
                df_existing = pd.read_excel(EXCEL_FILE_PATH_IN_PAGE)
                
                # Ensure all required columns exist
                for col in new_row_data.keys():
                    if col not in df_existing.columns:
                        df_existing[col] = None
            except Exception:
                # If file doesn't exist or is empty, create with all columns from new_row_data
                df_existing = pd.DataFrame(columns=list(new_row_data.keys()))
            
            # Create a new DataFrame with the new row
            new_row_df = pd.DataFrame([new_row_data])
            
            # Append the new row to the existing data
            df_updated = pd.concat([df_existing, new_row_df], ignore_index=True)
            
            # Save back to Excel (similar to CT9 approach)
            df_updated.to_excel(EXCEL_FILE_PATH_IN_PAGE, index=False)
                
            st.success("Données enregistrées avec succès!")
            st.success(f"Data successfully added and saved to: {EXCEL_FILE_PATH_IN_PAGE}")
            st.balloons()

        except Exception as e:
            st.error(f"Error saving data to Excel for sheet '{SHEET_NAME}': {e}")

    st.sidebar.markdown("--- ")
    if st.sidebar.checkbox(f"Show existing data?", key=f"show_data_{SHEET_NAME.replace(' ', '_')}"):
        try:
            df_display = pd.read_excel(EXCEL_FILE_PATH_IN_PAGE)
            st.sidebar.subheader(f"Current Direction Data (first 5 rows)")
            st.sidebar.dataframe(df_display.head())
        except Exception as e:
            st.sidebar.error(f"Could not display data: {e}")

except FileNotFoundError:
    st.error(f"ERREUR: Fichier Excel introuvable à l'emplacement {EXCEL_FILE_PATH_IN_PAGE}.")
except Exception as e:
    st.error(f"An unexpected error occurred: {e}")
    st.info("Please ensure the Excel file exists and is correctly formatted.")
