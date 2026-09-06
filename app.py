import streamlit as st
import pandas as pd
from io import BytesIO

st.title("🚒 Feuille de Garde")

@st.cache_data
def charger_gabarit(fichier):
    # Lecture avec encodage latin-1 pour les fichiers Windows/Excel
    with open(fichier, 'r', encoding='latin-1') as f:
        lignes = f.readlines()
        
    donnees = []
    agres = None
    fonctions_valides = ["CA", "COND", "EQ", "CE B1", "EQ B1", "CE B2", "EQ B2", "OBS", "COND/EQ"]
    mots_ignores = ["BILLET", "DE", "GARDE", "EQUIPE 2", "EQUIPE 3"]
    
    for ligne in lignes:
        # Uniformisation des séparateurs pour gérer les CSV français
        ligne = ligne.strip().replace(';', ',')
        
        # Ignorer les lignes vides
        if not ligne or ligne.replace(',', '').strip() == "": 
            continue
            
        # Nettoyage des espaces autour de chaque élément
        elements = [e.strip() for e in ligne.split(',')]
        
        # 1. Détection d'un agrès
        if elements[0] not in fonctions_valides and elements[0] not in mots_ignores and len(elements) > 1 and elements[1].isdigit():
            agres = elements[0]
            
        # 2. Détection d'une fonction sous l'agrès
        elif agres and elements[0] in fonctions_valides:
            personnel = elements[1] if len(elements) > 1 and elements[1] != "" else ""
            donnees.append({"Agrès": agres, "Fonction": elements[0], "Personnel": personnel})
            
    return pd.DataFrame(donnees)

# 1. Chargement du tableau
df = charger_gabarit("TEST_FEUILLE_DE_GARDE_BILLET.csv")

st.write("Affectation des équipages :")

# 2. Affichage du tableau modifiable en ligne
df_modifie = st.data_editor(df, num_rows="dynamic", use_container_width=True)

# 3. Création du fichier Excel en mémoire (C'est ici que 'output' est défini)
output = BytesIO()
with pd.ExcelWriter(output, engine='openpyxl') as writer:
    df_modifie.to_excel(writer, index=False, sheet_name='Garde')

# 4. Bouton de téléchargement
st.download_button(
    label="📥 Télécharger la feuille (Excel)", 
    data=output.getvalue(), 
    file_name="feuille_de_garde.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
