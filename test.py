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

# --- Interaction avec l'utilisateur ---
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
        st.subheader("📊 Aperçu du Dashboard basé sur ta sélection")

        # Mini sommaire cliquable
        st.markdown("### 📑 Sommaire des KPIs :")
        for idx, kpi in enumerate(st.session_state.kpis_valides, 1):
            st.markdown(f"- [{kpi.splitlines()[0]}](#kpi-{idx})")

        st.markdown("---")

        cols = st.columns(3)
        for idx, kpi in enumerate(st.session_state.kpis_valides):
            with cols[idx % 3]:
                st.markdown(f"<h4 id='kpi-{idx+1}'>KPI {idx+1}</h4>", unsafe_allow_html=True)
                st.markdown(kpi)
                st.divider()

    st.subheader("💬 Pose une nouvelle question à l'IA (améliorer / filtrer les KPIs) :")
    user_prompt = st.text_area("Ta question :")

    if st.button("Envoyer ma demande 🧠"):
        if user_prompt:
            save_prompt_to_airtable(user_prompt)
            historique = load_prompts_from_airtable()
            historique_contextuel = "\n".join(historique[-5:])

            prompt_final = f"""
Voici la liste actuelle des KPIs :
{chr(10).join(st.session_state.kpis)}

Voici l'historique récent des demandes :
{historique_contextuel}

Nouvelle demande de l'utilisateur :
{user_prompt}

Merci d'ajouter ou d'adapter des KPIs pertinents à la liste existante sans supprimer les anciens.
Ne te limite pas à 5 KPIs, ajoute autant que nécessaire selon la demande.
Pour chaque KPI ajouté :
- un titre clair
- une description
- un exemple de valeur ou formule
- un type de graphique adapté.
Retourne toute la liste complète des KPIs mise à jour.
"""
            try:
                response_update = client.chat.completions.create(
                    model="gpt-4",
                    messages=[{"role": "user", "content": prompt_final}],
                    temperature=0.5,
                    max_tokens=2000
                )
                st.session_state.kpis = response_update.choices[0].message.content.split("\n\n")
                st.success("✅ Liste de KPIs mise à jour avec succès !")

                st.subheader("📊 Liste actuelle des KPIs :")
                for idx, kpi in enumerate(st.session_state.kpis, 1):
                    st.markdown(f"**{idx}.** {kpi}")

            except Exception as e:
                st.error(f"❌ Une erreur est survenue : {e}")
