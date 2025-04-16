import streamlit as st
import pandas as pd
import openai
import io
import xlsxwriter

# 🧠 Paramètres GPT (remplace par ta clé)
openai.api_key = st.secrets.get("OPENAI_API_KEY", "sk-...")
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
        with st.spinner("Analyse du fichier avec LLM..."):
            prompt = f"""
Voici un extrait de données sous forme de tableau :

{df.head(10).to_markdown()}

Quels sont les KPIs intéressants à calculer ? Propose des types de graphiques pertinents.
"""
            try:
                response = openai.ChatCompletion.create(
                    model=GPT_MODEL,
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0.3
                )
                kpis = response.choices[0].message["content"]
            except:
                kpis = "Exemple (mode mock) :\n- Total des ventes\n- Ventes par produit\n- Évolution des ventes par mois"

        st.subheader("💡 Suggestions de KPIs")
        st.markdown(kpis)

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

