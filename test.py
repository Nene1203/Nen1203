import streamlit as st
import pandas as pd
import openai
import io
import xlsxwriter
import os
import requests
import datetime
from dotenv import load_dotenv

# --- Charger les variables d'environnement ou secrets ---
load_dotenv()

# --- Gestion du Thème (Dark / Light Mode) ---
if "theme" not in st.session_state:
    st.session_state.theme = "dark"  # Dark mode par défaut

theme_toggle = st.toggle("🌑 / ☀️ Changer de mode", value=(st.session_state.theme == "dark"))

if theme_toggle:
    st.session_state.theme = "dark"
else:
    st.session_state.theme = "light"

# --- Application du style selon le thème ---
if st.session_state.theme == "dark":
    st.markdown("""
    <style>
        body {
            background-color: #121212;
            color: #F1F1F1;
        }
        .stApp {
            background-color: #121212;
            color: #F1F1F1;
        }
        .css-1d391kg, .css-1cpxqw2 {
            background-color: #1F1F1F;
            border-radius: 10px;
            padding: 15px;
        }
        button {
            background-color: #2980b9;
            color: white;
        }
    </style>
    """, unsafe_allow_html=True)
else:
    st.markdown("""
    <style>
        body {
            background-color: #FFFFFF;
            color: #000000;
        }
        .stApp {
            background-color: #FFFFFF;
            color: #000000;
        }
        .css-1d391kg, .css-1cpxqw2 {
            background-color: #F9F9F9;
            border-radius: 10px;
            padding: 15px;
        }
        button {
            background-color: #3498db;
            color: white;
        }
    </style>
    """, unsafe_allow_html=True)

# --- Configuration des clés API ---
airtable_api_key = st.secrets["airtable"]["api_key"] if "airtable" in st.secrets else os.getenv("AIRTABLE_API_KEY")
airtable_base_id = st.secrets["airtable"]["base_id"] if "airtable" in st.secrets else os.getenv("AIRTABLE_BASE_ID")
airtable_table_name = st.secrets["airtable"].get("table_name", "Prompts") if "airtable" in st.secrets else os.getenv("AIRTABLE_TABLE_NAME")

openai.api_key = st.secrets.get("OPENAI_API_KEY") or os.getenv("OPENAI_API_KEY")
client = openai.OpenAI(api_key=openai.api_key)

# --- Fonctions Airtable ---
def save_prompt_to_airtable(prompt_text):
    url = f"https://api.airtable.com/v0/{airtable_base_id}/{airtable_table_name}"
    headers = {
        "Authorization": f"Bearer {airtable_api_key}",
        "Content-Type": "application/json"
    }
    data = {
        "fields": {
            "PromptText": prompt_text,
            "Timestamp": datetime.datetime.now().isoformat()
        }
    }
    response = requests.post(url, headers=headers, json=data)
    return response.status_code == 200

def load_prompts_from_airtable():
    url = f"https://api.airtable.com/v0/{airtable_base_id}/{airtable_table_name}?sort[0][field]=Timestamp&sort[0][direction]=asc"
    headers = {
        "Authorization": f"Bearer {airtable_api_key}",
        "Content-Type": "application/json"
    }
    response = requests.get(url, headers=headers)
    prompts = []
    if response.status_code == 200:
        records = response.json()["records"]
        for record in records:
            prompts.append(record["fields"].get("PromptText", ""))
    return prompts

# --- Interface utilisateur ---
st.title("POC SaaS : Générateur et Sélectionneur intelligent de KPIs 📊")

if "dashboard_ready" not in st.session_state:
    st.session_state.dashboard_ready = False

uploaded_file = st.file_uploader("Charge ton fichier Excel ou CSV ici", type=["xlsx", "csv"])

if uploaded_file:
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    st.subheader("Aperçu des données")
    st.dataframe(df.head(10))

    # --- Détection automatique du domaine ---
    columns = [col.lower() for col in df.columns]
    if any(x in columns for x in ["client", "produit", "vente", "chiffre d'affaires", "commande"]):
        domaine_detecte = "commercial"
    elif any(x in columns for x in ["email", "campagne", "clics", "impressions", "trafic"]):
        domaine_detecte = "marketing"
    elif any(x in columns for x in ["entrepôt", "stock", "livraison", "réception"]):
        domaine_detecte = "logistique"
    elif any(x in columns for x in ["patient", "date de naissance", "maladie", "décès"]):
        domaine_detecte = "santé"
    else:
        domaine_detecte = "général"

    st.subheader(f"🔎 Domaine détecté automatiquement : **{domaine_detecte.capitalize()}**")

    confirmation = st.radio("Est-ce correct ?", ("Oui", "Non"))

    if confirmation == "Oui":
        domaine = domaine_detecte
    else:
        domaine = st.selectbox(
            "Quel est le bon domaine ?",
            options=["commercial", "marketing", "logistique", "santé", "ressources humaines", "finance", "autre"]
        )

    if confirmation:
        contexte = {
            "commercial": "Données de ventes d'une entreprise commerciale.",
            "marketing": "Données de campagnes marketing et publicitaires.",
            "logistique": "Données de gestion de stocks, livraisons et entrepôts.",
            "santé": "Données médicales concernant des patients.",
            "ressources humaines": "Données de gestion RH et employés.",
            "finance": "Données financières et comptables.",
            "autre": "Données diverses sans domaine spécifique."
        }.get(domaine, "Données diverses sans domaine spécifique.")

        if st.button("Suggérer des KPIs 📈"):
            with st.spinner("Analyse des données et historique en cours..."):
                sample_data = df.sample(min(len(df), 20), random_state=42).to_csv(index=False)

                historique = load_prompts_from_airtable()
                historique_contextuel = "\n".join(historique[-5:])

                prompt_final = f"""
Voici un extrait de données :
{sample_data}

Contexte : {contexte}

Historique des questions précédentes :
{historique_contextuel}

Propose 5 KPIs pertinents pour ces données :
- un titre clair
- une description
- un exemple de valeur ou formule
- un type de graphique adapté
"""
                try:
                    response = client.chat.completions.create(
                        model="gpt-4",
                        messages=[{"role": "user", "content": prompt_final}],
                        temperature=0.5,
                        max_tokens=800
                    )
                    kpis = response.choices[0].message.content.split("\n\n")
                    st.session_state.kpis = kpis
                    st.success("✅ KPIs générés avec succès !")
                except Exception as e:
                    st.error(f"Erreur GPT : {e}")

# --- Interaction utilisateur ---
if "kpis" in st.session_state:
    st.subheader("✅ Sélectionne les KPIs pour ton Dashboard :")
    selected_kpis = []
    for kpi in st.session_state.kpis:
        if st.checkbox(kpi):
            selected_kpis.append(kpi)

    if selected_kpis:
        st.subheader("🚀 KPIs sélectionnés :")
        for kpi in selected_kpis:
            st.markdown(f"- {kpi}")

        if st.button("✅ Valider ma sélection de KPIs"):
            st.session_state.kpis_valides = selected_kpis
            st.success("✅ Sélection validée ! Prêt pour l'aperçu du Dashboard.")

if "kpis_valides" in st.session_state:
    if st.button("📊 Dashboard Preview"):
        st.session_state.dashboard_ready = True
        st.subheader("📑 Dashboard Preview élégant")

        st.markdown("### Sommaire 📚")
        for idx, kpi in enumerate(st.session_state.kpis_valides, 1):
            st.markdown(f"- [{kpi.splitlines()[0]}](#kpi-{idx})")

        st.markdown("---")

        cols = st.columns(3)
        for idx, kpi in enumerate(st.session_state.kpis_valides):
            with cols[idx % 3]:
                st.markdown(f"<div style='background-color: #f0f2f6; padding: 15px; border-radius: 10px; margin-bottom: 20px;'>"
                            f"<h4 id='kpi-{idx+1}' style='color: #2c3e50;'>📊 KPI {idx+1}</h4>"
                            f"<p style='font-size: 15px; color: #34495e;'>{kpi}</p>"
                            f"</div>", unsafe_allow_html=True)

    if st.session_state.dashboard_ready:
        st.subheader("📥 Exporte ton Dashboard Excel ici")

        with st.spinner('🔄 Génération de ton fichier Excel... Patiente quelques secondes...'):
            output = io.BytesIO()

            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='Données', index=False)
                workbook = writer.book

                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'middle',
                    'align': 'center',
                    'bg_color': '#DCE6F1',
                    'border': 1
                })

                cell_format = workbook.add_format({
                    'valign': 'middle',
                    'align': 'center',
                    'border': 1
                })

                for idx, kpi in enumerate(st.session_state.kpis_valides, start=1):
                    sheet_name = f"KPI_{idx}"
                    worksheet = workbook.add_worksheet(sheet_name)

                    worksheet.write('A1', df.columns[0], header_format)
                    worksheet.write('B1', df.columns[1], header_format)

                    for row in range(min(10, len(df))):
                        worksheet.write(row + 1, 0, str(df.iloc[row, 0]), cell_format)
                        worksheet.write(row + 1, 1, df.iloc[row, 1], cell_format)

                    worksheet.set_column('A:A', 20)
                    worksheet.set_column('B:B', 15)

                    chart = workbook.add_chart({'type': 'column'})
                    chart.add_series({
                        'name':       f'KPI {idx}',
                        'categories': f'={sheet_name}!$A$2:$A${min(11, len(df)+1)}',
                        'values':     f'={sheet_name}!$B$2:$B${min(11, len(df)+1)}',
                        'data_labels': {'value': True}
                    })
                    chart.set_title({'name': f'Dashboard KPI {idx}'})
                    chart.set_x_axis({'name': df.columns[0]})
                    chart.set_y_axis({'name': df.columns[1]})

                    worksheet.insert_chart('D2', chart)

        st.success("✅ Fichier Excel généré avec succès !")

        st.download_button(
            label="📥 Télécharger le Dashboard Excel Ultra-Pro",
            data=output.getvalue(),
            file_name="dashboard_kpis_ultra_premium.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
