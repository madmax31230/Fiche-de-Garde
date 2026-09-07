import streamlit as st
import pandas as pd
from io import BytesIO

# 1. Configuration de la page
st.set_page_config(page_title="Feuille de Garde - L'Isle-en-Dodon", page_icon="🚒", layout="wide")

# 2. Styles CSS personnalisés
st.markdown("""
    <style>
    .titre-caserne {
        color: #d32f2f;
        font-size: 2.5rem;
        font-weight: 800;
        margin-bottom: 0rem;
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
def charger_donnees_excel(fichier):
    # Chargement du gabarit de garde depuis l'onglet 'BILLET'
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
                
    df_garde = pd.DataFrame(donnees)

    # Chargement de l'onglet des véhicules s'il existe
    try:
        df_vehicules = pd.read_excel(fichier, sheet_name='VEHICULES')
    except Exception:
        # Valeurs par défaut si l'onglet n'est pas encore créé
        df_vehicules = pd.DataFrame(columns=["ENERGIE/TYPE", "NUMERO"])

    return df_garde, df_vehicules

fichier_source = "TEST FEUILLE DE GARDE.xlsx"

# En-tête de l'application
st.markdown('<p class="titre-caserne">🚒 Centre de Secours de L\'Isle-en-Dodon</p>', unsafe_allow_html=True)
st.markdown('<p class="sous-titre">Gestion opérationnelle des gardes et des véhicules</p>', unsafe_allow_html=True)

try:
    df_garde, df_vehicules = charger_donnees_excel(fichier_source)
    
    if df_garde.empty:
        st.warning("Le tableau de garde est vide. Vérifiez l'onglet 'BILLET'.")
    else:
        # Disposition en deux colonnes sur l'écran
        col_garde, col_vehicules = st.columns([2, 1])
        
        with col_garde:
            st.write("**📋 Affectation des équipages**")
            df_garde_modifie = st.data_editor(
                df_garde, 
                num_rows="dynamic", 
                use_container_width=True,
                height=550,
                hide_index=True 
            )
            
        with col_vehicules:
            st.write("**🚛 Gestion des Numéros de Véhicules**")
            df_vehicules_modifie = st.data_editor(
                df_vehicules, 
                num_rows="dynamic", 
                use_container_width=True,
                height=550,
                hide_index=True 
            )

        # --- PRÉPARATION DU FICHIER EXCEL DE SORTIE ---
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_garde_modifie.to_excel(writer, index=False, sheet_name='Garde')
            df_vehicules_modifie.to_excel(writer, index=False, sheet_name='VEHICULES')
        
        # --- BARRE LATÉRALE (SIDEBAR) ---
        effectif_saisi = df_garde_modifie['Personnel'].apply(lambda x: 1 if str(x).strip() != "" and str(x).lower() != "nan" else 0).sum()
        
        with st.sidebar:
            st.header("⚙️ Actions")
            st.metric(label="Postes pourvus", value=f"{effectif_saisi} / {len(df_garde)}")
            st.divider()
            st.write("Vérifiez les modifications avant l'exportation.")
            
            st.download_button(
                label="📥 Télécharger la feuille validée", 
                data=output.getvalue(), 
                file_name="Feuille_Garde_Modifiee.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary" 
            )

except Exception as e:
    st.error(f"⚠️ Erreur : {e}")
