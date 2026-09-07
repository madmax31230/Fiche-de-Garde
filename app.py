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

# Ton ID Google Sheets (récupéré depuis l'URL de ton tableur en ligne)
GOOGLE_SHEET_ID = "1WbCH8Q4r2rM2WL1f8KP2o7XaADi-1vjC"

@st.cache_data(ttl=60) # Actualisation automatique toutes les minutes
def charger_donnees_depuis_gsheets(sheet_id):
    try:
        url_billet = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet=BILLET"
        df_brut = pd.read_csv(url_billet, header=None)
    except Exception:
        df_brut = pd.DataFrame()

    donnees = []
    fonctions_valides = ["CA", "COND", "EQ", "CE B1", "EQ B1", "CE B2", "EQ B2", "OBS", "COND/EQ"]
    mots_ignores = ["BILLET", "DE", "GARDE", "EQUIPE", "LUNDI", "MARDI", "MERCREDI", 
                    "JEUDI", "VENDREDI", "SAMEDI", "DIMANCHE", "JANVIER", "FEVRIER", 
                    "MARS", "AVRIL", "MAI", "JUIN", "JUILLET", "AOUT", "SEPTEMBRE", 
                    "OCTOBRE", "NOVEMBRE", "DECEMBRE"]
    
    if not df_brut.empty:
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

    # Chargement de l'effectif directement depuis Google Sheets
    liste_agents = []
    dict_agents = {}
    try:
        url_effectif = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet=EFFECTIF"
        df_eff_brut = pd.read_csv(url_effectif, header=None)
        
        for r in range(1, len(df_eff_brut)):
            row = df_eff_brut.iloc[r]
            nom = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            prenom = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
            grade = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
            
            if not nom or nom.lower() == "nan":
                continue
                
            nom_simple = f"{nom} {prenom}"
            
            # Récupération de toutes les spécialités (colonnes F / index 5 et suivantes)
            specs = []
            for c in range(5, len(row)):
                val = row.iloc[c]
                if pd.notna(val) and str(val).strip() != "" and str(val).lower() != "nan":
                    specs.append(str(val).strip())
            
            details = []
            if grade and grade.lower() != 'nan':
                details.append(grade)
            if specs:
                details.extend(specs)
                
            if details:
                label = f"{nom_simple} ({' - '.join(details)})"
            else:
                label = nom_simple
                
            liste_agents.append(label)
            dict_agents[label] = nom_simple
            
        liste_agents = sorted(list(set(liste_agents)))
    except Exception:
        pass

    return df_garde, liste_agents, dict_agents

# En-tête
st.markdown('<p class="titre-caserne">🚒 Centre de Secours de L\'Isle-en-Dodon</p>', unsafe_allow_html=True)
st.markdown('<p class="sous-titre">Synchronisation en direct avec Google Sheets</p>', unsafe_allow_html=True)

try:
    df_garde, liste_agents, dict_agents = charger_donnees_depuis_gsheets(GOOGLE_SHEET_ID)
    
    if df_garde.empty:
        st.warning("Impossible de lire l'onglet 'BILLET' de votre Google Sheets. Vérifiez que le tableur est bien partagé ('Accès général : Tous les utilisateurs ayant le lien').")
    else:
        st.write("### 📋 Équipe de garde du jour (Modifiable en temps réel)")
        
        lignes_mises_a_jour = []
        
        for idx, row in df_garde.iterrows():
            col1, col2, col3 = st.columns([1.5, 1, 3])
            with col1:
                st.markdown(f"**{row['Agrès']}**")
            with col2:
                st.markdown(f"`{row['Fonction']}`")
            with col3:
                agent_actuel = row['Personnel']
                if liste_agents:
                    options = [""] + liste_agents
                    default_idx = 0
                    for i, opt in enumerate(options):
                        if agent_actuel.strip().lower() in opt.lower():
                            default_idx = i
                            break
                            
                    choix_label = st.selectbox(
                        f"Agent {idx}", 
                        options=options, 
                        index=default_idx, 
                        label_visibility="collapsed",
                        key=f"agent_{idx}"
                    )
                    nouveau_personnel = dict_agents.get(choix_label, choix_label.split(" (")[0] if choix_label else "")
                else:
                    nouveau_personnel = st.text_input(
                        f"Agent {idx}", 
                        value=agent_actuel, 
                        label_visibility="collapsed",
                        key=f"text_agent_{idx}"
                    )
            
            lignes_mises_a_jour.append({
                "Agrès": row['Agrès'],
                "Fonction": row['Fonction'],
                "Personnel": nouveau_personnel
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
    st.error(f"⚠️ Erreur de connexion au Google Sheets : {e}")
