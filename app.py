import streamlit as st
import pandas as pd
from io import BytesIO

st.title("🚒 Feuille de Garde")

@st.cache_data
def charger_gabarit(fichier):
    with open(fichier, 'r', encoding='utf-8') as f:
        lignes = f.readlines()
    donnees = []
    agres = None
    for ligne in lignes:
        ligne = ligne.strip()
        if not ligne or ligne == ",,": continue
        elements = ligne.split(',')
        
        if elements[0] not in ["CA", "COND", "EQ", "CE B1", "EQ B1", "CE B2", "EQ B2", "OBS", "BILLET", "DE", "GARDE", "EQUIPE 2"] and len(elements) > 1 and elements[1].isdigit():
            agres = elements[0]
        elif agres and elements[0] in ["CA", "COND", "EQ", "CE B1", "EQ B1", "CE B2", "EQ B2", "OBS"]:
            donnees.append({"Agrès": agres, "Fonction": elements[0], "Personnel": ""})
    return pd.DataFrame(donnees)

df = charger_gabarit("TEST_FEUILLE_DE_GARDE_BILLET.csv")

st.write("Affectation des équipages :")
# Tableau modifiable directement sur la page web
df_modifie = st.data_editor(df, num_rows="dynamic", use_container_width=True)

# Export Excel
output = BytesIO()
with pd.ExcelWriter(output, engine='openpyxl') as writer:
    df_modifie.to_excel(writer, index=False, sheet_name='Garde')

st.download_button(label="📥 Télécharger la feuille (Excel)", data=output.getvalue(), file_name="feuille_de_garde.xlsx")