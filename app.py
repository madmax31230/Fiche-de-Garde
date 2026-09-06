import streamlit as st
import pandas as pd
from io import BytesIO

st.title("🚒 Feuille de Garde")

@st.cache_data
def charger_gabarit(fichier):
    # Remplacer 'utf-8' par 'latin-1' sur cette ligne :
    with open(fichier, 'r', encoding='latin-1') as f:
        lignes = f.readlines()
df = charger_gabarit("TEST_FEUILLE_DE_GARDE_BILLET.csv")

st.write("Affectation des équipages :")
# Tableau modifiable directement sur la page web
df_modifie = st.data_editor(df, num_rows="dynamic", use_container_width=True)

# Export Excel
output = BytesIO()
with pd.ExcelWriter(output, engine='openpyxl') as writer:
    df_modifie.to_excel(writer, index=False, sheet_name='Garde')

st.download_button(label="📥 Télécharger la feuille (Excel)", data=output.getvalue(), file_name="feuille_de_garde.xlsx")
