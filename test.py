import streamlit as st
import pandas as pd
from openai import OpenAI
import io
import xlsxwriter

# 🧠 Paramètres GPT (remplace par ta clé)
OpenAI.api_key = st.secrets.get("OPENAI_API_KEY", "sk-...")
GPT_MODEL = "gpt-4"

st.title("POC SaaS : Générateur de Dashboard Excel Automatisé 📊")

# 📥 Upload du fichier
uploaded_file = st.file_uploader("Charge ton fichier Excel ici", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.subheader("Aperçu des données")
    st.dataframe(df.head(20))

    # ✅ Appel LLM ou mock
    if st.button("Suggérer des KPIs 📈"):
        with st.spinner("Analyse des données en cours... 🤖"):
            sample_data = df.head(10).to_csv(index=False)

            prompt = f"""Voici un extrait de données au format CSV :

{sample_data}

Propose 5 KPIs pertinents à calculer à partir de ces données.
Pour chaque KPI, donne :
- un titre clair,
- une description,
- un exemple de valeur ou de formule,
- un type de graphique adapté (barres, camembert, lignes, histogramme, etc.)."""

        try:
            response = openai.ChatCompletion.create(
                model="gpt-4",
                messages=[
                    {"role": "system", "content": "Tu es un expert en BI et dashboards interactifs."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.7,
                max_tokens=800
            )

            kpis = response["choices"][0]["message"]["content"]
            st.success("✅ Analyse terminée")
            st.markdown("### 📊 KPIs suggérés par GPT-4 :")
            st.markdown(kpis)

        except Exception as e:
            st.error(f"❌ Une erreur est survenue : {e}")
    # 📊 Génération fichier Excel
    if st.button("Générer fichier Excel avec Dashboard 🔄"):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, sheet_name="Données", index=False)
            workbook = writer.book
            worksheet = writer.sheets["Données"]

            # Exemple de graphique
            chart = workbook.add_chart({"type": "column"})
            chart.add_series({
                "categories": f"=Données!$A$2:$A$10",
                "values":     f"=Données!$B$2:$B$10",
                "name":       "Exemple"
            })
            worksheet.insert_chart("E5", chart)

            # Ajouter le script Python dans un onglet
            code_sheet = workbook.add_worksheet("Code_Python")
            code_content = '''import pandas as pd\n# Exemple de traitement\n'''
            for i, line in enumerate(code_content.split("\n")):
                code_sheet.write(i, 0, line)

        st.success("Fichier généré !")
        st.download_button(
            label="📥 Télécharger le fichier Excel",
            data=output.getvalue(),
            file_name="dashboard_generé.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

