import streamlit as st
import pandas as pd
from io import BytesIO

# 1. Configuration de la page
st.set_page_config(page_title="Feuille de Garde - L'Isle-en-Dodon", page_icon="🚒", layout="wide")

# 2. Injection de CSS personnalisé
st.markdown("""
    <style>
    .titre-caserne {
        color: #d32f2f;
        font-size: 2.5rem;
        font-weight: 800;
        margin-bottom: 0rem;
        padding-bottom: 0rem;
    }
    .sous-titre {
        color: #666;
        font-size: 1.2rem;
        font-style: italic;
        margin-bottom: 2rem;
    }
    </style>
""", unsafe_allow_html=True)

@st.cache_data
def charger_gabarit_excel(fichier):
    df_brut = pd.read_excel(fichier, sheet_name='BILLET', header=None)
    donnees = []
    fonctions_valides = ["CA", "COND", "EQ", "CE B1", "EQ B1", "CE B2", "EQ B2", "OBS", "COND/EQ"]
    mots_ignores = ["BILLET", "DE", "GARDE", "EQUIPE", "LUNDI", "MARDI", "MERCREDI", 
                    "JEUDI", "VENDREDI", "SAMEDI", "DIMANCHE", "JANVIER", "FEVRIER", 
                    "MARS", "AVRIL", "MAI", "JUIN", "JUILLET", "AOUT", "SEPTEMBRE", 
                    "OCTOBRE", "NOVEMBRE", "DECEMBRE"]
    
    for col_idx in range(df_brut.shape[1]):
        agres_en_cours = None
        for row_idx in range(df_brut.shape[0]):
            valeur = str(df_brut.iloc[row_idx, col_idx]).strip()
            
            if pd.isna(df_brut.iloc[row_idx, col_idx]) or valeur.lower() == "nan" or valeur == "":
                continue
                
            valeur_maj = valeur.upper()
            is_fonction = False
            
            for f in fonctions_valides:
                if valeur_maj == f:
                    is_fonction = True
                    fonction_trouvee = f
                    break
                    
            is_ignore = any(ignore in valeur_maj for ignore in mots_ignores)
            
            if is_fonction:
                if agres_en_cours:
                    personnel = ""
                    if col_idx + 1 < df_brut.shape[1]:
                        p = str(df_brut.iloc[row_idx, col_idx + 1]).strip()
                        if p.lower() != "nan" and p != "" and p.upper() not in fonctions_valides:
                            personnel = p
                    donnees.append({"Agrès": agres_en_cours, "Fonction": fonction_trouvee, "Personnel": personnel})
            
            elif not is_ignore and len(valeur_maj) >= 2:
                agres_en_cours = valeur
                
    return pd.DataFrame(donnees)

fichier_source = "TEST FEUILLE DE GARDE.xlsx"

# En-tête
st.markdown('<p class="titre-caserne">🚒 Centre de Secours de L\'Isle-en-Dodon</p>', unsafe_allow_html=True)
st.markdown('<p class="sous-titre">Gestion opérationnelle de la feuille de garde</p>', unsafe_allow_html=True)

try:
    df = charger_gabarit_excel(fichier_source)
    
    if df.empty:
        st.warning("Le tableau est vide. Vérifiez que l'onglet s'appelle bien 'BILLET'.")
    else:
        # --- ZONE PRINCIPALE ---
        st.write("**📋 Affectation des équipages**")
        df_modifie = st.data_editor(
            df, 
            num_rows="dynamic", 
            use_container_width=True,
            height=600,
            hide_index=True 
        )
        
        # --- PRÉPARATION DU FICHIER EXCEL (Généré après l'édition) ---
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_modifie.to_excel(writer, index=False, sheet_name='Garde')
        
        # --- BARRE LATÉRALE ---
        effectif_saisi = df_modifie['Personnel'].apply(lambda x: 1 if str(x).strip() != "" and str(x).lower() != "nan" else 0).sum()
        
        with st.sidebar:
            st.header("⚙️ Actions")
            st.metric(label="Agents affectés", value=f"{effectif_saisi} / {len(df)}")
            st.divider()
            st.write("Vérifiez les affectations avant de générer le fichier Excel.")
            
            st.download_button(
                label="📥 Télécharger la feuille validée", 
                data=output.getvalue(), 
                file_name="Feuille_Garde_Modifiee.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary" 
            )

except Exception as e:
    st.error(f"⚠️ Erreur de lecture du fichier Excel : {e}")
