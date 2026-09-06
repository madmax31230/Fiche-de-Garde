import streamlit as st
import pandas as pd
from io import BytesIO

st.title("🚒 Feuille de Garde")

@st.cache_data
def charger_gabarit(fichier):
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
        
        # 1. Détection d'un agrès (si l'élément 1 est un chiffre, ex: VSAV, 17)
        if elements[0] not in fonctions_valides and elements[0] not in mots_ignores and len(elements) > 1 and elements[1].isdigit():
            agres = elements[0]
            
        # 2. Détection d'une fonction sous l'agrès en cours
        elif agres and elements[0] in fonctions_valides:
            # Récupérer un nom si déjà présent dans le CSV d'origine
            personnel = elements[1] if len(elements) > 1 and elements[1] != "" else ""
            donnees.append({"Agrès": agres, "Fonction": elements[0], "Personnel": personnel})
            
    return pd.DataFrame(donnees)
st.download_button(label="📥 Télécharger la feuille (Excel)", data=output.getvalue(), file_name="feuille_de_garde.xlsx")
