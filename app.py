import streamlit as st
import pandas as pd
from io import BytesIO

st.title("🚒 Feuille de Garde")

@st.cache_data
def charger_gabarit_excel(fichier):
    # Lecture directe du fichier Excel, en ciblant spécifiquement l'onglet 'BILLET'
    df_brut = pd.read_excel(fichier, sheet_name='BILLET', header=None)
    
    donnees = []
    agres = None
    fonctions_valides = ["CA", "COND", "EQ", "CE B1", "EQ B1", "CE B2", "EQ B2", "OBS", "COND/EQ"]
    mots_ignores = ["BILLET", "DE", "GARDE", "EQUIPE 2", "EQUIPE 3"]
    
    # On parcourt chaque ligne du fichier Excel
    for index, row in df_brut.iterrows():
        # On récupère le texte de la première cellule de la ligne
        valeur_cellule = str(row[0]).strip()
        
        # Ignorer les lignes vides ("nan" est la façon dont pandas lit une case vide)
        if valeur_cellule == "nan" or valeur_cellule == "":
            continue
            
        # Si on a écrit "VSAV,17" dans une seule case Excel, on sépare le nom
        elements = [e.strip() for e in valeur_cellule.replace(';', ',').split(',')]
        mot_principal = elements[0].upper()
        
        # 1. Détection de l'engin (si ce n'est pas une fonction et pas un mot ignoré)
        if mot_principal not in fonctions_valides and mot_principal not in mots_ignores:
            agres = elements[0]  # On garde le nom complet (ex: VSAV ou VSRM)
            
        # 2. Détection d'une fonction sous l'engin
        elif agres and mot_principal in fonctions_valides:
            personnel = elements[1] if len(elements) > 1 else ""
            donnees.append({"Agrès": agres, "Fonction": mot_principal, "Personnel": personnel})
            
    return pd.DataFrame(donnees)

fichier_source = "TEST FEUILLE DE GARDE.xlsx"

try:
    # 1. Chargement du tableau depuis le fichier Excel
    df = charger_gabarit_excel(fichier_source)
    st.write("Affectation des équipages :")
    
    # 2. Affichage du tableau modifiable en ligne
    df_modifie = st.data_editor(df, num_rows="dynamic", use_container_width=True)
    
    # 3. Préparation du téléchargement
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_modifie.to_excel(writer, index=False, sheet_name='Garde')

    # 4. Bouton de téléchargement
    st.download_button(
        label="📥 Télécharger la feuille validée (Excel)", 
        data=output.getvalue(), 
        file_name="Feuille_Garde_Modifiee.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

except FileNotFoundError:
    st.error(f"⚠️ Le fichier {fichier_source} est introuvable sur GitHub.")
except ValueError:
    st.error("⚠️ L'onglet 'BILLET' est introuvable dans votre fichier Excel. Vérifiez le nom de l'onglet.")
except Exception as e:
    st.error(f"⚠️ Une erreur s'est produite : {e}")
