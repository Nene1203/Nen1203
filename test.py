# --- Imports ---
import streamlit as st
import pandas as pd
import openai
import io
import os
import requests
import datetime
import xlwings as xw
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

def generate_vba_code_from_kpis(selected_kpis):
    vba_code = "Sub CreerDashboard()\n\n"
    vba_code += "    Dim wsData As Worksheet\n"
    vba_code += "    Dim wsPivot As Worksheet\n"
    vba_code += "    Dim pvtCache As PivotCache\n"
    vba_code += "    Dim pvt As PivotTable\n"
    vba_code += "    Dim chartObj As ChartObject\n\n"
    vba_code += "    Set wsData = ThisWorkbook.Sheets(\"Données\")\n"
    vba_code += "    Set wsPivot = ThisWorkbook.Sheets.Add\n"
    vba_code += "    wsPivot.Name = \"Dashboard\"\n\n"
    vba_code += "    Set pvtCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=wsData.UsedRange)\n\n"

    for idx, kpi in enumerate(selected_kpis, 1):
        pivot_name = f"TCD_KPI_{idx}"
        chart_title = f"KPI {idx} - {kpi.splitlines()[0][:30]}"
        vba_code += f"    ' --- {chart_title} ---\n"
        vba_code += f"    Set pvt = pvtCache.CreatePivotTable(TableDestination:=wsPivot.Cells({3 + idx * 15}, 1), TableName:=\"{pivot_name}\")\n"
        vba_code += f"    Set chartObj = wsPivot.ChartObjects.Add(Left:=300, Width:=400, Top:={50 + idx * 300}, Height:=250)\n"
        vba_code += f"    With chartObj.Chart\n"
        vba_code += f"        .SetSourceData Source:=pvt.TableRange2\n"
        vba_code += f"        .ChartType = xlColumnClustered\n"
        vba_code += f"        .HasTitle = True\n"
        vba_code += f"        .ChartTitle.Text = \"{chart_title}\"\n"
        vba_code += f"    End With\n\n"

    vba_code += "End Sub\n"
    return vba_code

def create_xlsm_dashboard(df, selected_kpis, output_xlsm_path):
    temp_xlsx = output_xlsm_path.replace('.xlsm', '.xlsx')
    df.to_excel(temp_xlsx, sheet_name='Données', index=False)
    app = xw.App(visible=False)
    wb = app.books.open(temp_xlsx)
    wb.save(output_xlsm_path)
    vba_code = generate_vba_code_from_kpis(selected_kpis)
    wb.api.VBProject.VBComponents.Add(1).CodeModule.AddFromString(vba_code)
    wb.api.VBProject.VBComponents("ThisWorkbook").CodeModule.AddFromString("""
Private Sub Workbook_Open()
    Call CreerDashboard
End Sub
""")
    wb.save()
    wb.close()
    app.quit()
    os.remove(temp_xlsx)

# --- OpenAI Client ---
openai.api_key = st.secrets.get("OPENAI_API_KEY") or os.getenv("OPENAI_API_KEY")
client = openai.OpenAI(api_key=openai.api_key)

# --- Airtable config ---
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

# --- Pages principales ---
def page_accueil_premium():
    st.title("🚀 SaaS KPI Generator")
    st.write("Automatisez vos dashboards comme jamais auparavant.")
    st.image("https://images.unsplash.com/photo-1612832021092-6cc8fb0b5fb3?ixlib=rb-4.0.3&auto=format&fit=crop&w=1350&q=80", use_container_width=True)

def page_generation_dashboard():
    uploaded_file = st.file_uploader("📂 Déposez votre fichier Excel (.xlsx) ou CSV :", type=["xlsx", "csv"])
    if uploaded_file:
        df = pd.read_excel(uploaded_file) if uploaded_file.name.endswith('.xlsx') else pd.read_csv(uploaded_file)
        st.subheader("Aperçu des données")
        st.dataframe(df.head(10))

        if st.button("Suggérer des KPIs"):
            with st.spinner("Analyse intelligente..."):
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
                st.success("✅ KPIs générés !")

        if "kpis" in st.session_state:
            st.subheader("Sélectionne les KPIs pour ton Dashboard")
            selected_kpis = []
            for kpi in st.session_state.kpis:
                if st.checkbox(kpi):
                    selected_kpis.append(kpi)

            if selected_kpis and st.button("Exporter mon Dashboard"):
                output_path = "dashboard_auto.xlsx"
                create_xlsm_dashboard(df, selected_kpis, output_path)
                save_project_to_airtable("Projet_" + datetime.datetime.now().strftime("%Y%m%d_%H%M"), output_path)
                with open(output_path, "rb") as file:
                    st.download_button(
                        label="📥 Télécharger Dashboard Excel",
                        data=file,
                        file_name="dashboard_kpi_ultra.xlsm",
                        mime="application/vnd.ms-excel.sheet.macroEnabled.12"
                    )

def page_historique_projets():
    st.title("📚 Historique des projets créés")
    historique = load_projects_from_airtable()
    if not historique.empty:
        st.dataframe(historique)
    else:
        st.info("Aucun projet sauvegardé pour le moment.")

# --- Barre de navigation ---
page = st.sidebar.selectbox(
    "Navigation",
    ("🏠 Accueil", "🛠️ Générer Dashboard", "📚 Historique des Projets")
)

# --- Routing ---
if page == "🏠 Accueil":
    page_accueil_premium()
elif page == "🛠️ Générer Dashboard":
    page_generation_dashboard()
elif page == "📚 Historique des Projets":
    page_historique_projets()

# --- Footer ---
footer_premium()
