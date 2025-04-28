import streamlit as st
import pandas as pd
import openai
import io
import os
import requests
import datetime
import xlwings as xw
from dotenv import load_dotenv

# 🚨 Cette ligne doit être tout en haut (obligatoire Streamlit)
st.set_page_config(page_title="SaaS KPI Generator", page_icon="🚀", layout="wide")

# --- Animation subtile pour page d'accueil ---
st.markdown("""
    <style>
        @keyframes zoomFade {
            0% {transform: scale(0.9); opacity: 0;}
            100% {transform: scale(1); opacity: 1;}
        }
        h1 {
            animation: zoomFade 1s ease-in-out;
        }
    </style>
""", unsafe_allow_html=True)

# --- Pied de page premium ---
def footer_premium():
    st.markdown("""
        <hr style="margin-top:50px;margin-bottom:10px;">
        <div style="text-align: center; color: gray;">
            <p>© 2025 Nelson Telep - Tous droits réservés 🚀</p>
        </div>
    """, unsafe_allow_html=True)

# --- Charger les variables d'environnement ou secrets ---
load_dotenv()

# --- Page d'accueil Premium ---
def page_accueil_premium():
    st.markdown("""
        <div style='text-align: center; margin-bottom: 50px;'>
            <h1 style='font-size: 50px; color: #3498db;'>🚀 SaaS KPI Generator</h1>
            <p style='font-size: 20px; color: gray;'>Automatisez vos dashboards comme jamais auparavant</p>
        </div>
    """, unsafe_allow_html=True)

    st.image("https://images.unsplash.com/photo-1612832021092-6cc8fb0b5fb3?ixlib=rb-4.0.3&auto=format&fit=crop&w=1350&q=80", use_column_width=True)

    st.markdown("""
        <div style='margin-top: 50px;'>
            <h2 style='color: #2c3e50;'>Pourquoi choisir notre solution ?</h2>
            <ul style='font-size:18px;'>
                <li>🚀 Générez automatiquement vos KPIs stratégiques</li>
                <li>📊 Créez des dashboards Excel ultra professionnels sans coder</li>
                <li>🤖 Personnalisez vos indicateurs grâce à l'intelligence artificielle</li>
            </ul>
        </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("<h3 style='text-align:center;'>Commencez maintenant 🚀</h3>", unsafe_allow_html=True)
    st.markdown("<p style='text-align:center;'>Chargez votre premier fichier Excel ou CSV pour découvrir la magie ✨</p>", unsafe_allow_html=True)

# --- Gestion du Thème (Dark / Light Mode) ---
if "theme" not in st.session_state:
    st.session_state.theme = "dark"

theme_toggle = st.toggle("🌑 / ☀️ Changer de mode", value=(st.session_state.theme == "dark"))

if theme_toggle:
    st.session_state.theme = "dark"
else:
    st.session_state.theme = "light"

if st.session_state.theme == "dark":
    st.markdown("""
    <style>
        body {background-color: #121212; color: #F1F1F1;}
        .stApp {background-color: #121212; color: #F1F1F1;}
        .css-1d391kg, .css-1cpxqw2 {background-color: #1F1F1F; border-radius: 10px; padding: 15px;}
        button {background-color: #2980b9; color: white;}
    </style>
    """, unsafe_allow_html=True)
else:
    st.markdown("""
    <style>
        body {background-color: #FFFFFF; color: #000000;}
        .stApp {background-color: #FFFFFF; color: #000000;}
        .css-1d391kg, .css-1cpxqw2 {background-color: #F9F9F9; border-radius: 10px; padding: 15px;}
        button {background-color: #3498db; color: white;}
    </style>
    """, unsafe_allow_html=True)

# --- Configuration OpenAI et Airtable ---
airtable_api_key = st.secrets["airtable"]["api_key"] if "airtable" in st.secrets else os.getenv("AIRTABLE_API_KEY")
airtable_base_id = st.secrets["airtable"]["base_id"] if "airtable" in st.secrets else os.getenv("AIRTABLE_BASE_ID")
airtable_table_name = st.secrets["airtable"].get("table_name", "Prompts") if "airtable" in st.secrets else os.getenv("AIRTABLE_TABLE_NAME")

openai.api_key = st.secrets.get("OPENAI_API_KEY") or os.getenv("OPENAI_API_KEY")
client = openai.OpenAI(api_key=openai.api_key)

# --- Fonctions Airtable ---
def save_prompt_to_airtable(prompt_text):
    url = f"https://api.airtable.com/v0/{airtable_base_id}/{airtable_table_name}"
    headers = {"Authorization": f"Bearer {airtable_api_key}", "Content-Type": "application/json"}
    data = {"fields": {"PromptText": prompt_text, "Timestamp": datetime.datetime.now().isoformat()}}
    response = requests.post(url, headers=headers, json=data)
    return response.status_code == 200

def load_prompts_from_airtable():
    url = f"https://api.airtable.com/v0/{airtable_base_id}/{airtable_table_name}?sort[0][field]=Timestamp&sort[0][direction]=asc"
    headers = {"Authorization": f"Bearer {airtable_api_key}", "Content-Type": "application/json"}
    response = requests.get(url, headers=headers)
    prompts = []
    if response.status_code == 200:
        records = response.json()["records"]
        for record in records:
            prompts.append(record["fields"].get("PromptText", ""))
    return prompts

# --- Fonctions Excel VBA ---
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

    open_macro = """
Private Sub Workbook_Open()
    Call CreerDashboard
End Sub
"""
    wb.api.VBProject.VBComponents("ThisWorkbook").CodeModule.AddFromString(open_macro)

    wb.save()
    wb.close()
    app.quit()

    os.remove(temp_xlsx)

# --- Interface ---
page_accueil_premium()

uploaded_file = st.file_uploader("📂 Déposez un fichier Excel (.xlsx) ou CSV ici :", type=["xlsx", "csv"])

# À continuer ensuite avec l'upload, l'analyse, la génération de KPIs, l'export en XLSM...
