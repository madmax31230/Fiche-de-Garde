import streamlit as st
import pandas as pd
from io import BytesIO

# 1. Configuration de la page
st.set_page_config(page_title="Feuille de Garde - L'Isle-en-Dodon", page_icon="🚒", layout="wide")

# 2. Design et styles
st.markdown("""
    <style>
    .titre-caserne {
        color: #d32f2f;
        font-size: 2.2rem;
        font-weight: 800;
        margin-bottom: 0rem;
    }
    .sous-titre {
        color: #666;
        font-size: 1.1rem;
        font-style: italic;
        margin-bottom: 1.5rem;
    }
    </style>
""", unsafe_allow_html=True)

@st.cache_data
def charger_donnees(fichier):
    # Chargement de la garde (onglet BILLET)
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

    # Chargement de l'effectif pour les listes déroulantes
    try:
        df_effectif = pd.read_excel(fichier, sheet_name='EFFECTIF')
        # On crée une colonne Nom complet propre (ex: Dupont Jean)
        if 'NOM' in df_effectif.columns and 'PRENOM' in df_effectif.columns:
            df_effectif['AGENT'] = df_effectif['NOM'].astype(str) + " " + df_effectif['PRENOM'].astype(str)
            liste_agents = sorted(df_effectif['AGENT'].dropna().unique().tolist())
        else:
            liste_agents = []
    except Exception:
        liste_agents = []

    return df_garde, liste_agents

fichier_source = "feuille_garde.xlsx"

# En-tête
st.markdown('<p class="titre-caserne">🚒 Centre de Secours de L\'Isle-en-Dodon</p>', unsafe_allow_html=True)
st.markdown('<p class="sous-titre">Gestion opérationnelle de l\'équipe de garde</p>', unsafe_allow_html=True)

try:
    df_garde, liste_agents = charger_donnees(fichier_source)
    
    if df_garde.empty:
        st.warning("Le tableau de garde est vide. Vérifiez l'onglet 'BILLET'.")
    else:
        st.write("### 📋 Équipe de garde du jour (Modifiable)")
        
        # On prépare une interface claire pour modifier les agents via des listes déroulantes
        lignes_mises_a_jour = []
        
        # Pour chaque poste, on propose un affichage en ligne propre avec un selectbox si des agents sont dispos
        for idx, row in df_garde.iterrows():
            col1, col2, col3 = st.columns([1.5, 1, 2.5])
            with col1:
                st.markdown(f"**{row['Agrès']}**")
            with col2:
                st.markdown(f"`{row['Fonction']}`")
            with col3:
                agent_actuel = row['Personnel']
                # Si la liste d'effectif existe, on propose un selectbox, sinon un champ texte
                if liste_agents:
                    options = [""] + liste_agents
                    default_idx = options.index(agent_actuel) if agent_actuel in options else 0
                    nouveau_ PERSONNEL = st.selectbox(
                        f"Agent {idx}", 
                        options=options, 
                        index=default_idx, 
                        label_visibility="collapsed",
                        key=f"agent_{idx}"
                    )
                else:
                    nouveau_ PERSONNEL = st.text_input(
                        f"Agent {idx}", 
                        value=agent_actuel, 
                        label_visibility="collapsed",
                        key=f"text_agent_{idx}"
                    )
            
            lignes_mises_a_jour.append({
                "Agrès": row['Agrès'],
                "Fonction": row['Fonction'],
                "Personnel": nouveau__ PERSONNEL if 'nouveau__ PERSONNEL' in locals() else nouveau_ PERSONNEL
            })

        df_final = pd.DataFrame(lignes_mises_a_jour)

        # --- EXPORT ET ACTIONS ---
        st.divider()
        col_g, col_d = st.columns([2, 1])
        with col_d:
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Garde')
            
            st.download_button(
                label="📥 Télécharger la feuille validée (Excel)", 
                data=output.getvalue(), 
                file_name="Feuille_Garde_Mise_A_Jour.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

except Exception as e:
    st.error(f"⚠️ Erreur : {e}")
