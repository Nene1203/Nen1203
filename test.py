# --- Imports ---
import streamlit as st
import pandas as pd
import openai
import io
import os
import requests
import datetime
from dotenv import load_dotenv
import xlwings as xw
import shutil

# --- Config Streamlit ---
st.set_page_config(page_title="SaaS KPI Generator Premium", page_icon="🚀", layout="wide")

# --- Chargement des variables d'environnement ---
load_dotenv()

openai.api_key = st.secrets.get("OPENAI_API_KEY") or os.getenv("OPENAI_API_KEY")
client = openai.OpenAI(api_key=openai.api_key)

# --- Fonctions Génération VBA ---
def generate_vba_code_from_kpis(selected_kpis):
    vba_code = "Sub CreerDashboard()\n\n"
    vba_code += "    Dim wsData As Worksheet\n"
    vba_code += "    Dim wsDashboard As Worksheet\n"
    vba_code += "    Dim pvtCache As PivotCache\n"
    vba_code += "    Dim pvt As PivotTable\n"
    vba_code += "    Dim chartObj As ChartObject\n"
    vba_code += "    Dim lastRow As Long\n\n"
    vba_code += "    Set wsData = ThisWorkbook.Sheets(\"Données\")\n"
    vba_code += "    Set wsDashboard = ThisWorkbook.Sheets.Add\n"
    vba_code += "    wsDashboard.Name = \"Dashboard\"\n\n"
    vba_code += "    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row\n\n"

    top_position = 20
    left_position = 400

    for idx, kpi in enumerate(selected_kpis, 1):
        title = kpi.splitlines()[0] if kpi else f"KPI {idx}"

        if "produit" in title.lower():
            dimension = "Produit"
            chart_type = "xlColumnClustered"
        elif "région" in title.lower():
            dimension = "Région"
            chart_type = "xlPie"
        elif "mois" in title.lower() or "date" in title.lower():
            dimension = "Date"
            chart_type = "xlLine"
        else:
            dimension = "Produit"
            chart_type = "xlColumnClustered"

        pivot_name = f"TCD_KPI_{idx}"

        vba_code += f"    ' --- {title} ---\n"
        vba_code += f"    Set pvtCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=wsData.Name & \"!A1:D\" & lastRow)\n"
        vba_code += f"    Set pvt = pvtCache.CreatePivotTable(TableDestination:=wsDashboard.Cells({2 + idx * 15}, 2), TableName:=\"{pivot_name}\")\n"
        vba_code += f"    With pvt\n"
        vba_code += f"        .PivotFields(\"{dimension}\").Orientation = xlRowField\n"
        vba_code += f"        .PivotFields(\"{dimension}\").Position = 1\n"
        vba_code += f"        .PivotFields(\"Ventes\").Orientation = xlDataField\n"
        vba_code += f"        .PivotFields(\"Ventes\").Function = xlSum\n"
        vba_code += f"    End With\n\n"

        vba_code += f"    Set chartObj = wsDashboard.ChartObjects.Add(Left:={left_position}, Width:=400, Top:={top_position}, Height:=300)\n"
        vba_code += f"    With chartObj.Chart\n"
        vba_code += f"        .SetSourceData Source:=pvt.TableRange2\n"
        vba_code += f"        .ChartType = {chart_type}\n"
        vba_code += f"        .HasTitle = True\n"
        vba_code += f"        .ChartTitle.Text = \"{title}\"\n"
        vba_code += f"    End With\n\n"

        top_position += 350

    vba_code += "End Sub\n"
    return vba_code

# --- Fonction Création XLSM ---
def create_xlsm_with_vba(df, selected_kpis, output_path="dashboard_final.xlsm"):
    temp_xlsx = output_path.replace('.xlsm', '.xlsx')
    df.to_excel(temp_xlsx, sheet_name='Données', index=False)

    app = xw.App(visible=False)
    wb = app.books.open(temp_xlsx)

    wb.save(output_path)

    vba_code = generate_vba_code_from_kpis(selected_kpis)

    wb.api.VBProject.VBComponents.Add(1).CodeModule.AddFromString(vba_code)

    thisworkbook_code = wb.api.VBProject.VBComponents("ThisWorkbook").CodeModule
    thisworkbook_code.AddFromString("""
Private Sub Workbook_Open()
    Call CreerDashboard
End Sub
""")

    wb.save()
    wb.close()
    app.quit()

    os.remove(temp_xlsx)

    return output_path

# --- Interface Streamlit ---
st.title("🚀 SaaS KPI Generator - Version Premium Automatique")

uploaded_file = st.file_uploader("Chargez votre fichier Excel (.xlsx) ou CSV :", type=["xlsx", "csv"])

if uploaded_file:
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    st.subheader("🔎 Aperçu rapide")
    st.dataframe(df.head(10))

    if st.button("✨ Suggérer des KPIs avec GPT-4"):
        with st.spinner("Analyse intelligente en cours..."):
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
    st.subheader("📊 Sélectionnez vos KPIs préférés")

    selected_kpis = []
    for kpi in st.session_state.kpis:
        if st.checkbox(kpi):
            selected_kpis.append(kpi)

    if selected_kpis and st.button("📥 Exporter Dashboard Excel Premium (.xlsm)"):
        output_path = create_xlsm_with_vba(df, selected_kpis)
        with open(output_path, "rb") as f:
            st.download_button(
                label="📥 Télécharger votre Dashboard Ultra Premium",
                data=f,
                file_name="dashboard_kpis_premium.xlsm",
                mime="application/vnd.ms-excel.sheet.macroEnabled.12"
            )

# Footer Premium
st.markdown("""
---
<div style='text-align: center; color: gray;'>
    © 2025 Nelson Telep - Powered by AI 🚀
</div>
""", unsafe_allow_html=True)
