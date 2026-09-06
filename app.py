import streamlit as st
import pandas as pd
from io import BytesIO

st.title("🚒 Feuille de Garde")

@st.cache_data
def charger_gabarit_excel(fichier):
    df_brut = pd.read_excel(fichier, sheet_name='BILLET', header=None)
    donnees = []
    fonctions_valides = ["CA", "COND", "EQ", "CE B1", "EQ B1", "CE B2", "EQ B2", "OBS", "COND/EQ"]
    
    mots_ignores = ["BILLET", "DE", "GARDE", "EQUIPE", "LUNDI", "MARDI", "MERCREDI", 
                    "JEUDI", "VENDREDI", "SAMEDI", "DIMANCHE", "JANVIER", "FEVRIER", 
                    "MARS", "AVRIL", "MAI", "JUIN", "JUILLET", "AOUT", "SEPTEMBRE", 
                    "OCTOBRE", "NOVEMBRE", "DECEMBRE"]
    
    # Lecture verticale : colonne par colonne
    for col_idx in range(df_brut.shape[1]):
        agres_en_cours = None
        
        for row_idx in range(df_brut.shape[0]):
            valeur = str(df_brut.iloc[row_idx, col_idx]).strip()
            
            # Ignorer les cellules vides
            if pd.isna(df_brut.iloc[row_idx, col_idx]) or valeur.lower() == "nan" or valeur == "":
                continue
                
            valeur_maj = valeur.upper()
            
            # Est-ce une fonction ?
            is_fonction = False
            for f in fonctions_valides:
                if valeur_maj == f:
                    is_fonction = True
                    fonction_trouvee = f
                    break
                    
            is_ignore = any(ignore in valeur_maj for ignore in mots_ignores)
            
            if is_fonction:
                if agres_en_cours:
                    # Le personnel est cherché dans la cellule juste à droite
                    personnel = ""
                    if col_idx + 1 < df_brut.shape[1]:
                        p = str(df_brut.iloc[row_idx, col_idx + 1]).strip()
                        if p.lower() != "nan" and p != "" and p.upper() not in fonctions_valides:
                            personnel = p
                            
                    donnees.append({"Agrès": agres_en_cours, "Fonction": fonction_trouvee, "Personnel": personnel})
            
            # Si ce n'est ni une fonction ni un mot ignoré, c'est un nouvel engin
            elif not is_ignore and len(valeur_maj) >= 2:
                agres_en_cours = valeur
                
    return pd.DataFrame(donnees)

fichier_source = "TEST_FEUILLE_DE_GARDE.xlsx"

try:
    df = charger_gabarit_excel(fichier_source)
    
    if df.empty:
        st.warning("Le tableau est toujours vide. Vérifiez que l'onglet s'appelle bien 'BILLET'.")
    else:
        st.write("Affectation des équipages :")
        df_modifie = st.data_editor(df, num_rows="dynamic", use_container_width=True)
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_modifie.to_excel(writer, index=False, sheet_name='Garde')

        st.download_button(
            label="📥 Télécharger la feuille validée (Excel)", 
            data=output.getvalue(), 
            file_name="Feuille_Garde_Modifiee.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

except Exception as e:
    st.error(f"⚠️ Erreur de lecture du fichier Excel : {e}")
