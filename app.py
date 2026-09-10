import streamlit as st
import pandas as pd
from io import BytesIO

# 1. Configuration de la page en mode large
st.set_page_config(
    page_title="Feuille de Garde - L'Isle-en-Dodon", 
    page_icon="🚒", 
    layout="wide"
)

# 2. Styles CSS pour un rendu propre et structuré
st.markdown("""
    <style>
    .main {
        background-color: #0e1117;
    }
    .header-box {
        background: linear-gradient(90deg, #b71c1c 0%, #d32f2f 100%);
        padding: 15px;
        border-radius: 8px;
        color: white;
        text-align: center;
        margin-bottom: 20px;
    }
    .header-title {
        font-size: 1.8rem;
        font-weight: 800;
        margin: 0;
    }
    .header-subtitle {
        font-size: 1rem;
        font-style: italic;
        margin: 5px 0 0 0;
        opacity: 0.9;
    }
    .section-title {
        color: #ff5252;
        font-size: 1.3rem;
        font-weight: 700;
        margin-top: 25px;
        margin-bottom: 10px;
        border-bottom: 2px solid #333;
        padding-bottom: 5px;
    }
    </style>
""", unsafe_allow_html=True)

# Ton ID Google Sheets
GOOGLE_SHEET_ID = "1WbCH8Q4r2rM2WL1f8KP2o7XaADi-1vjC"

@st.cache_data(ttl=60)
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
                    "OCTOBRE", "NOVEMBRE", "DECEMBRE", "CONSIGNES", "SPORT", "FMA"]
    
    nom_engin_actuel = "GENERAL"
    if not df_brut.empty:
        for col_idx in range(df_brut.shape[1]):
            for row_idx in range(df_brut.shape[0]):
                valeur = str(df_brut.iloc[row_idx, col_idx]).strip()
                if pd.isna(df_brut.iloc[row_idx, col_idx]) or valeur.lower() == "nan" or valeur == "":
                    continue
                valeur_maj = valeur.upper()
                
                is_fonction = valeur_maj in fonctions_valides
                is_ignore = any(ignore in valeur_maj for ignore in mots_ignores)
                
                if is_fonction:
                    personnel = ""
                    if col_idx + 1 < df_brut.shape[1]:
                        p = str(df_brut.iloc[row_idx, col_idx + 1]).strip()
                        if p.lower() != "nan" and p != "" and p.upper() not in fonctions_valides:
                            personnel = p
                    donnees.append({"Agrès": nom_engin_actuel, "Fonction": valeur_maj, "Personnel": personnel})
                elif not is_ignore and len(valeur_maj) >= 2 and not is_fonction:
                    num_indicatif = ""
                    if col_idx + 1 < df_brut.shape[1]:
                        val_suiv = str(df_brut.iloc[row_idx, col_idx + 1]).strip()
                        if val_suiv.replace('.', '', 1).isdigit():
                            num_indicatif = str(int(float(val_suiv)))
                    elif row_idx + 1 < df_brut.shape[0]:
                        val_dessous = str(df_brut.iloc[row_idx + 1, col_idx]).strip()
                        if val_dessous.replace('.', '', 1).isdigit():
                            num_indicatif = str(int(float(val_dessous)))
                            
                    if num_indicatif:
                        nom_engin_actuel = f"{valeur} {num_indicatif}"
                    else:
                        nom_engin_actuel = valeur
                    
    df_garde = pd.DataFrame(donnees)

    # Chargement de l'effectif depuis Google Sheets avec toutes les spécialités
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

# En-tête visuel
st.markdown("""
    <div class="header-box">
        <p class="header-title">🚒 CENTRE DE SECOURS DE L'ISLE-EN-DODON</p>
        <p class="header-subtitle">Feuille de Garde - Tri personnalisé par taille d'équipage</p>
    </div>
""", unsafe_allow_html=True)

try:
    df_garde, liste_agents, dict_agents = charger_donnees_depuis_gsheets(GOOGLE_SHEET_ID)
    
    if df_garde.empty:
        st.warning("⚠️ Impossible de lire l'onglet 'BILLET' de votre Google Sheets.")
    else:
        agres_uniques = df_garde['Agrès'].unique()
        groupes_par_taille = {}
        
        for agres in agres_uniques:
            df_agres = df_garde[df_garde['Agrès'] == agres]
            nb_postes = len(df_agres)
            if nb_postes not in groupes_par_taille:
                groupes_par_taille[nb_postes] = []
            groupes_par_taille[nb_postes].append((agres, df_agres))

        lignes_mises_a_jour = []

        # --- ORDRE Souhaité DES TAILLES D'ÉQUIPAGES ---
        # Tu peux modifier cet ordre directement ici (ex: [4, 6, 3, 2, 5])
        ordre_tailles_souhaite = [4, 6, 3, 2, 5]

        for nb_postes in ordre_tailles_souhaite:
            if nb_postes in groupes_par_taille:
                st.markdown(f'<div class="section-title">Équipages à {nb_postes} postes</div>', unsafe_allow_html=True)
                
                vehicules_du_groupe = groupes_par_taille[nb_postes]
                
                for i_veh in range(0, len(vehicules_du_groupe), 3):
                    cols_ligne = st.columns(3)
                    batch = vehicules_du_groupe[i_veh:i_veh+3]
                    
                    for idx_col, (agres, df_agres) in enumerate(batch):
                        with cols_ligne[idx_col]:
                            st.markdown(f"### 🚚 {agres}")
                            for i, row in df_agres.iterrows():
                                cols_poste = st.columns([1, 2.5])
                                with cols_poste[0]:
                                    st.markdown(f"`{row['Fonction']}`")
                                with cols_poste[1]:
                                    agent_actuel = row['Personnel']
                                    options = [""] + liste_agents if liste_agents else [""]
                                    default_idx = 0
                                    for opt_idx, opt in enumerate(options):
                                        if agent_actuel.strip().lower() in opt.lower():
                                            default_idx = opt_idx
                                            break
                                            
                                    choix_label = st.selectbox(
                                        f"{agres}_{row['Fonction']}_{i}", 
                                        options=options, 
                                        index=default_idx, 
                                        label_visibility="collapsed",
                                        key=f"agent_{i}_{agres}"
                                    )
                                    nouveau_personnel = dict_agents.get(choix_label, choix_label.split(" (")[0] if choix_label else "")
                                
                                lignes_mises_a_jour.append({
                                    "Agrès": agres,
                                    "Fonction": row['Fonction'],
                                    "Personnel": nouveau_personnel
                                })
                            st.divider()

        with st.sidebar:
            st.markdown("### 📥 Actions")
            output = BytesIO()
            df_final = pd.DataFrame(lignes_mises_a_jour)
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Garde')
            
            st.download_button(
                label="📥 Télécharger la Feuille Validée", 
                data=output.getvalue(), 
                file_name="Feuille_Garde_Mise_A_Jour.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )

except Exception as e:
    st.error(f"⚠️ Erreur : {e}")
