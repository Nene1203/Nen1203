# --- Imports ---
import streamlit as st
import pandas as pd
import openai
import io
import os
import requests
import datetime
from dotenv import load_dotenv

# --- Config Streamlit ---
st.set_page_config(page_title="SaaS KPI Generator", page_icon="🚀", layout="wide")

# --- Chargement des variables d'environnement ---
load_dotenv()

# --- Fonctions auxiliaires ---
def footer_premium():
    st.markdown("""
        <hr style="margin-top:50px;margin-bottom:10px;">
        <div style="text-align: center; color: gray;">
            <p>© 2025 Nelson Telep - Tous droits réservés 🚀</p>
        </div>
    """, unsafe_allow_html=True)

# --- Nouvelle fonction : création Dashboard Excel (.xlsx) sans VBA ---
def create_xlsx_dashboard(df, selected_kpis):
    import xlsxwriter

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Onglet Données
        df.to_excel(writer, sheet_name='Données', index=False)
        workbook = writer.book

        # Onglet Résumé KPIs
        worksheet_summary = workbook.add_worksheet("KPIs Résumés")

        # Ecrire les KPIs
        worksheet_summary.write('A1', 'KPI')
        worksheet_summary.write('B1', 'Description')

        for idx, kpi in enumerate(selected_kpis, start=2):
            title = kpi.splitlines()[0] if kpi else f"KPI {idx-1}"
            description = "Basé sur vos données chargées."
            worksheet_summary.write(f"A{idx}", title)
            worksheet_summary.write(f"B{idx}", description)

        worksheet_summary.set_column('A:A', 40)
        worksheet_summary.set_column('B:B', 60)

        # Créer un graphique
        worksheet_chart = workbook.add_worksheet("Graphiques")
        chart = workbook.add_chart({'type': 'column'})
        chart.add_series({
            'categories': '=KPIs Résumés!$A$2:$A$6',
            'values': '=KPIs Résumés!$B$2:$B$6',
            'name': 'KPIs',
            'data_labels': {'value': True}
        })
        chart.set_title({'name': 'Distribution des KPIs'})
        chart.set_x_axis({'name': 'Indicateurs'})
        chart.set_y_axis({'name': 'Scores fictifs'})
        worksheet_chart.insert_chart('B2', chart)

    output.seek(0)
    return output

# --- OpenAI Client ---
openai.api_key = st.secrets.get("OPENAI_API_KEY") or os.getenv("OPENAI_API_KEY")
client = openai.OpenAI(api_key=openai.api_key)

# --- Airtable Config ---
airtable_api_key = st.secrets["airtable"]["api_key"] if "airtable" in st.secrets else os.getenv("AIRTABLE_API_KEY")
airtable_base_id = st.secrets["airtable"]["base_id"] if "airtable" in st.secrets else os.getenv("AIRTABLE_BASE_ID")
airtable_table_name = st.secrets["airtable"].get("table_name", "Projets") if "airtable" in st.secrets else os.getenv("AIRTABLE_TABLE_NAME")

def save_project_to_airtable(project_name, file_name):
    url = f"https://api.airtable.com/v0/{airtable_base_id}/{airtable_table_name}"
    headers = {"Authorization": f"Bearer {airtable_api_key}", "Content-Type": "application/json"}
    data = {"fields": {"ProjectName": project_name, "DateCreated": datetime.datetime.now().isoformat(), "FileName": file_name}}
    requests.post(url, headers=headers, json=data)

def load_projects_from_airtable():
    url = f"https://api.airtable.com/v0/{airtable_base_id}/{airtable_table_name}?sort[0][field]=DateCreated&sort[0][direction]=desc"
    headers = {"Authorization": f"Bearer {airtable_api_key}", "Content-Type": "application/json"}
    response = requests.get(url, headers=headers)
    projects = []
    if response.status_code == 200:
        records = response.json()["records"]
        for record in records:
            fields = record["fields"]
            projects.append({"Nom du projet": fields.get("ProjectName", ""), "Date": fields.get("DateCreated", ""), "Fichier": fields.get("FileName", "")})
    return pd.DataFrame(projects)

# --- Pages ---
def page_accueil_et_generation():
    st.title("🚀 SaaS KPI Generator")
    st.write("Automatisez vos dashboards comme jamais auparavant.")
    st.image("https://images.unsplash.com/photo-1612832021092-6cc8fb0b5fb3?ixlib=rb-4.0.3&auto=format&fit=crop&w=1350&q=80", use_container_width=True)

    st.header("📂 Chargez votre fichier")
    uploaded_file = st.file_uploader("Déposez un Excel (.xlsx) ou CSV :", type=["xlsx", "csv"])
    
    if uploaded_file:
        df = pd.read_excel(uploaded_file) if uploaded_file.name.endswith('.xlsx') else pd.read_csv(uploaded_file)
        st.subheader("🔎 Aperçu des données")
        st.dataframe(df.head(10))

        if st.button("✨ Suggérer des KPIs"):
            with st.spinner("Analyse IA en cours..."):
                sample_data = df.sample(min(len(df), 20)).to_csv(index=False)
                prompt = f"""Voici un échantillon de données :\n{sample_data}\nPropose 5 KPIs pertinents."""
                response = client.chat.completions.create(
                    model="gpt-4",
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0.5,
                    max_tokens=800
                )
                kpis = response.choices[0].message.content.split("\n\n")
                st.session_state.kpis = kpis
                st.success("✅ KPIs générés avec succès !")

        if "kpis" in st.session_state:
            st.subheader("📊 Sélectionnez vos KPIs")
            selected_kpis = []
            for kpi in st.session_state.kpis:
                if st.checkbox(kpi):
                    selected_kpis.append(kpi)

            if selected_kpis and st.button("📥 Exporter Dashboard Excel (.xlsx)"):
                excel_file = create_xlsx_dashboard(df, selected_kpis)
                save_project_to_airtable("Projet_" + datetime.datetime.now().strftime("%Y%m%d_%H%M"), "dashboard_kpis_auto.xlsx")
                st.download_button(
                    label="📥 Télécharger le Dashboard Excel",
                    data=excel_file,
                    file_name="dashboard_kpis_auto.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

def page_historique_projets():
    st.title("📚 Historique des projets créés")
    historique = load_projects_from_airtable()
    if not historique.empty:
        st.dataframe(historique)
    else:
        st.info("Aucun projet sauvegardé pour le moment.")

# --- Navigation ---
page = st.sidebar.selectbox(
    "Navigation",
    ("🏠 Accueil & Génération", "📚 Historique des Projets")
)

# --- Routing ---
if page == "🏠 Accueil & Génération":
    page_accueil_et_generation()
elif page == "📚 Historique des Projets":
    page_historique_projets()

# --- Footer ---
footer_premium()
