import streamlit as st
import pandas as pd
import os
import plotly.express as px
from datetime import datetime
import os



# Set page configuration
st.set_page_config(
    page_title="Main Dashboard",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Dashboard Administratif Global")

import time

if st.button("🔄 Rafraîchir les données"):
    st.rerun()

# Show last reload time
st.caption(f"Dernier rechargement: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")

st.markdown("""
<style>
.big-metric {
    font-size: 2.5rem;
    font-weight: bold;
}
</style>
""", unsafe_allow_html=True)

DATA_FOLDER = "data"
# Création automatique du dossier s'il n'existe pas
if not os.path.exists(DATA_FOLDER):
    os.makedirs(DATA_FOLDER)

# Get all Excel files in data folder (ignore __pycache__ and non-xlsx)
excel_files = [f for f in os.listdir(DATA_FOLDER) if f.endswith('.xlsx') and not f.startswith('__')]

if not excel_files:
    st.error("Aucun fichier Excel trouvé dans le dossier 'data'.")
    st.stop()

# Create a tab for each Excel file
tabs = st.tabs([os.path.splitext(f)[0] for f in excel_files])

for i, excel_file in enumerate(excel_files):
    excel_path = os.path.join(DATA_FOLDER, excel_file)
    with tabs[i]:
        st.header(f"📄 Données: {excel_file}")
        try:
            df = pd.read_excel(excel_path)
            if df.empty:
                st.warning("Aucune donnée trouvée dans ce fichier.")
                continue

            # Show global metrics if relevant columns exist
            metric_cols = []
            if 'Réaliser' in df.columns:
                total_realiser = df['Réaliser'].sum()
                metric_cols.append(("Total Réalisé", total_realiser))
            if '%' in df.columns:
                avg_percent = df['%'].mean()
                metric_cols.append(("% Moyen Réalisation", f"{avg_percent:.2f}%"))
            if 'Project' in df.columns:
                n_projects = df['Project'].nunique()
                metric_cols.append(("Nombre de Projets", n_projects))
            if 'Operateurs Present' in df.columns:
                total_present = df['Operateurs Present'].sum()
                metric_cols.append(("Total Opérateurs Présents", total_present))
            if 'Production' in df.columns:
                total_prod = df['Production'].sum()
                metric_cols.append(("Production Totale", total_prod))
            if 'Maintenance_Total_Minutes_Ouverture' in df.columns:
                total_maint = df['Maintenance_Total_Minutes_Ouverture'].sum()
                metric_cols.append(("Maintenance (Total min Ouverture)", total_maint))
            if 'Visite_Date' in df.columns:
                n_visites = df['Visite_Date'].count()
                metric_cols.append(("Nb Visites Client/Autre", n_visites))

            if metric_cols:
                st.subheader("🔢 Indicateurs Clés")
                cols = st.columns(len(metric_cols))
                for j, (label, val) in enumerate(metric_cols):
                    with cols[j]:
                        st.markdown(f"<div class='big-metric'>{val}</div>", unsafe_allow_html=True)
                        st.caption(label)

            # Show a sample of the data
            st.subheader("🗂️ Aperçu des Données")
            st.dataframe(df.head(30), use_container_width=True)

            # If weekly columns exist, show a chart
            days = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi']
            if all(day in df.columns for day in days):
                st.subheader("📅 Production Hebdomadaire")
                weekly_data = df[days].sum()
                fig = px.bar(
                    weekly_data,
                    title="Production par Jour de la Semaine",
                    labels={'value': 'Quantité', 'index': 'Jour'},
                    color_discrete_sequence=['#4F8BF9']
                )
                st.plotly_chart(fig, use_container_width=True, key=f'{excel_file}_weekly_chart')

            # If Project exists, show project performance
            if 'Project' in df.columns and 'Réaliser' in df.columns and 'Objecti Semaine' in df.columns:
                st.subheader("📊 Performance par Projet")
                proj_perf = df.groupby('Project')[['Réaliser', 'Objecti Semaine']].sum()
                proj_perf['% Réalisation'] = (proj_perf['Réaliser'] / proj_perf['Objecti Semaine'] * 100).fillna(0)
                fig = px.bar(
                    proj_perf,
                    y='% Réalisation',
                    title="% Réalisation par Projet",
                    labels={'index': 'Projet', 'value': '% Réalisation'},
                    color_discrete_sequence=['#2E8B57']
                )
                st.plotly_chart(fig, use_container_width=True, key=f'{excel_file}_project_chart')

            # If Maintenance data exists, show maintenance summary
            if 'Maintenance_Total_Minutes_Ouverture' in df.columns:
                st.subheader("🛠️ Synthèse Maintenance")
                maint_cols = [col for col in df.columns if col.startswith('Maintenance_') and 'Total' not in col]
                if maint_cols:
                    st.dataframe(df[maint_cols + [c for c in df.columns if 'Total' in c and c.startswith('Maintenance_')]].head(30), use_container_width=True)

            # If Visite data exists, show visits table
            if 'Visite_Date' in df.columns:
                st.subheader("👥 Visites Client ou Autre")
                st.dataframe(df[['Visite_Date', 'Visite_Semaine', 'Visite_Motif', 'Visite_Qui']].dropna(how='all'), use_container_width=True)

        except Exception as e:
            st.error(f"Erreur lors du chargement de {excel_file} : {e}")
