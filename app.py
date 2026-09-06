import streamlit as st
import pandas as pd
from io import BytesIO

st.title("🚒 Feuille de Garde")

@st.cache_data
def charger_gabarit_excel(fichier):
    # Lecture de l'onglet BILLET sans se soucier des colonnes exactes
    df_brut = pd.read_excel(fichier, sheet_name='BILLET', header=None)
    
    donnees = []
    agres = None
    fonctions_valides = ["CA", "COND", "EQ", "CE B1", "EQ B1", "CE B2", "EQ B2", "OBS", "COND/EQ"]
    
    # Mots à ignorer pour ne pas les confondre avec des engins
    mots_ignores = ["BILLET", "DE", "GARDE", "EQUIPE", "LUNDI", "MARDI", "MERCREDI", 
                    "JEUDI", "VENDREDI", "SAMEDI", "DIMANCHE", "JANVIER", "FEVRIER", 
                    "MARS", "AVRIL", "MAI", "JUIN", "JUILLET", "AOUT", "SEPTEMBRE", 
                    "OCTOBRE", "NOVEMBRE", "DECEMBRE"]
    
    for index, row in df_brut.iterrows():
        # On récupère toutes les cellules non vides de la ligne, de gauche à droite
        valeurs = [str(v).strip() for v in row.values if pd.notna(v) and str(v).strip() != ""]
        
        # Si la ligne est complètement vide, on passe
        if not valeurs:
            continue
            
        premiere_valeur = valeurs[0].upper()
        
        # Est-ce une fonction ? (CA, COND...)
        is_fonction = False
        fonction_trouvee = ""
        for f in fonctions_valides:
            if premiere_valeur == f:
                is_fonction = True
                fonction_trouvee = f
                break
                
        # Est-ce un mot à ignorer (titre, date...) ?
        is_ignore = any(ignore in premiere_valeur for ignore in mots_ignores)
        
        # Si c'est une fonction et qu'on a déjà trouvé un engin au-dessus
        if is_fonction:
            if agres is not None:
                # Si un nom est déjà écrit à côté de la fonction, on le récupère
                personnel = valeurs[1] if len(valeurs) > 1 else ""
                donnees.append({"Agrès": agres, "Fonction": fonction_trouvee, "Personnel": personnel})
                
        # Sinon, si ce n'est pas ignoré et que c'est un mot d'au moins 2 lettres (ex: VL, FPT, VSAV), c'est un engin
        elif not is_ignore and len(premiere_valeur) >= 2:
            agres = valeurs[0] # On garde le texte exact (ex: "EPC - 13" ou "VSAV")
            
    return pd.DataFrame(donnees)

fichier_source = "TEST_FEUILLE_DE_GARDE.xlsx"

try:
    df = charger_gabarit_excel(fichier_source)
    
    if df.empty:
        st.warning("Le tableau est toujours vide. Vérifiez que l'onglet s'appelle bien 'BILLET' et contient vos données.")
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
