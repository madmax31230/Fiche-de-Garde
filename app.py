import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl
import os

st.set_page_config(
    page_title="Feuille de Garde - CS CARSALADE", 
    page_icon="🚒", 
    layout="wide"
)

st.markdown("""
    <style>
    .main { background-color: #0e1117; }
    .header-box {
        background: linear-gradient(90deg, #b71c1c 0%, #d32f2f 100%);
        padding: 15px; border-radius: 8px; color: white;
        text-align: center; margin-bottom: 20px;
    }
    .header-title { font-size: 1.8rem; font-weight: 800; margin: 0; }
    .header-subtitle { font-size: 1rem; font-style: italic; margin: 5px 0 0 0; opacity: 0.9; }
    .section-title {
        color: #ff5252; font-size: 1.3rem; font-weight: 700;
        margin-top: 25px; margin-bottom: 10px;
        border-bottom: 2px solid #333; padding-bottom: 5px;
    }
    .vsav-title {
        color: #4fc3f7;
        font-size: 1.5rem; font-weight: 800;
        margin-top: 10px; margin-bottom: 10px;
        border-bottom: 3px solid #0288d1; padding-bottom: 5px;
    }
    </style>
""", unsafe_allow_html=True)

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
                    "OCTOBRE", "NOVEMBRE", "DECEMBRE", "CONSIGNES", "SPORT", "FMA",
                    "SPECIALITE", "SPECIALITÉ", "SPÉCIALITÉ", "SPÉCIALITE"]
    
    INDICATIFS_AUTO = {
        "VSRM": "VSRM 10",
        "VFCDG": "VFCDG 88",
        "FPT": "FPT 11",
        "EPC": "EPC 13",
        "VID": "VID 34",
        "CCFM": "CCFM 29",
        "VSMPM": "VSMPM 02"
    }

    nom_engin_actuel = "GENERAL"
    vehicules_vus = {}
    anchor_csv = None

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
                    if anchor_csv is None:
                        anchor_csv = (row_idx, col_idx, valeur_maj)
                        
                    personnel = ""
                    if col_idx + 1 < df_brut.shape[1]:
                        p = str(df_brut.iloc[row_idx, col_idx + 1]).strip()
                        if p.lower() != "nan" and p != "" and p.upper() not in fonctions_valides:
                            personnel = p
                    donnees.append({
                        "row_idx": row_idx, 
                        "col_personnel": col_idx + 1, 
                        "Agrès": nom_engin_actuel, 
                        "Fonction": valeur_maj, 
                        "Personnel": personnel
                    })
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
                            
                    base_name = f"{valeur_maj} {num_indicatif}" if num_indicatif else valeur_maj
                    
                    if base_name == "VSAV":
                        if base_name in vehicules_vus:
                            vehicules_vus[base_name] += 1
                            if vehicules_vus[base_name] == 2:
                                nom_engin_actuel = "VSAV 17"
                            else:
                                nom_engin_actuel = f"VSAV {vehicules_vus[base_name]}"
                        else:
                            vehicules_vus[base_name] = 1
                            nom_engin_actuel = "VSAV 98"
                    elif base_name in INDICATIFS_AUTO:
                        nom_engin_actuel = INDICATIFS_AUTO[base_name]
                    else:
                        if base_name in vehicules_vus:
                            vehicules_vus[base_name] += 1
                            nom_engin_actuel = f"{base_name} {vehicules_vus[base_name]}"
                            if vehicules_vus[base_name] == 2:
                                for d in donnees:
                                    if d["Agrès"] == base_name:
                                        d["Agrès"] = f"{base_name} 1"
                        else:
                            vehicules_vus[base_name] = 1
                            nom_engin_actuel = base_name
                    
    df_garde = pd.DataFrame(donnees)

    liste_agents = []
    dict_agents = {}
    dict_specs = {}

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
                    specs.append(str(val).strip().upper())
            
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
            dict_specs[label] = specs
            
        liste_agents = sorted(list(set(liste_agents)))
    except Exception:
        pass

    return df_garde, liste_agents, dict_agents, anchor_csv, dict_specs

try:
    df_garde, liste_agents, dict_agents, anchor_csv, dict_specs = charger_donnees_depuis_gsheets(GOOGLE_SHEET_ID)
    
    if df_garde.empty:
        st.warning("⚠️ Impossible de lire l'onglet 'BILLET' de votre Google Sheets.")
    else:
        
        # --- FILTRE DE COMPÉTENCES INTELLIGENT ---
        def get_options_filtrees(agres_nom, fonction, agent_actuel_str):
            options_valides = [""]
            for opt in liste_agents:
                specs_agent = dict_specs.get(opt, [])
                est_autorise = True
                
                # Normalisation : supprime les espaces et tirets pour éviter les erreurs de syntaxe ("CA 1E" -> "CA1E")
                specs_clean = [str(s).replace(" ", "").replace("-", "") for s in specs_agent]
                
                # RÈGLE 1 : CA du VSAV (Accepte CA, CA1E, CATE...)
                if "VSAV" in agres_nom.upper() and fonction.upper() == "CA":
                    if not any(s.startswith("CA") for s in specs_clean):
                        est_autorise = False
                        
                # RÈGLE 2 : CA du FPT (Accepte uniquement CATE)
                elif "FPT" in agres_nom.upper() and fonction.upper() == "CA":
                    if not any("CATE" in s for s in specs_clean):
                        est_autorise = False
                        
                # L'agent est ajouté s'il est qualifié OU s'il était déjà inscrit par erreur dans le tableau source
                if est_autorise or (agent_actuel_str.strip() != "" and agent_actuel_str.strip().lower() in opt.lower()):
                    options_valides.append(opt)
                    
            return options_valides

        agres_uniques = df_garde['Agrès'].unique()
        modifications_agents = {}
        
        vsav_uniques = [a for a in agres_uniques if "VSAV" in a.upper()]
        autres_uniques = [a for a in agres_uniques if "VSAV" not in a.upper()]

        if vsav_uniques:
            st.markdown('<div class="vsav-title">🚑 VÉHICULES DE SECOURS AUX VICTIMES (VSAV)</div>', unsafe_allow_html=True)
            for i_veh in range(0, len(vsav_uniques), 3):
                cols_ligne = st.columns(3)
                batch = vsav_uniques[i_veh:i_veh+3]
                for idx_col, agres in enumerate(batch):
                    df_agres = df_garde[df_garde['Agrès'] == agres]
                    with cols_ligne[idx_col]:
                        st.markdown(f"### 🚚 {agres}")
                        for i, row in df_agres.iterrows():
                            cols_poste = st.columns([1, 2.5])
                            with cols_poste[0]:
                                st.markdown(f"`{row['Fonction']}`")
                            with cols_poste[1]:
                                agent_actuel = row['Personnel']
                                
                                options = get_options_filtrees(agres, row['Fonction'], agent_actuel)
                                
                                default_idx = 0
                                for opt_idx, opt in enumerate(options):
                                    if agent_actuel.strip().lower() in opt.lower() and agent_actuel.strip() != "":
                                        default_idx = opt_idx
                                        break
                                choix_label = st.selectbox(
                                    f"{agres}_{row['Fonction']}_{i}", 
                                    options=options, index=default_idx, 
                                    label_visibility="collapsed", key=f"agent_{i}_{agres}"
                                )
                                nouveau_personnel = dict_agents.get(choix_label, choix_label.split(" (")[0] if choix_label else "")
                            modifications_agents[(row['row_idx'], row['col_personnel'])] = nouveau_personnel
                        st.divider()

        groupes_par_taille = {}
        for agres in autres_uniques:
            df_agres = df_garde[df_garde['Agrès'] == agres]
            nb_postes = len(df_agres)
            if nb_postes not in groupes_par_taille:
                groupes_par_taille[nb_postes] = []
            groupes_par_taille[nb_postes].append((agres, df_agres))

        ordre_tailles_souhaite = [4, 6, 3, 2, 5]
        for taille in sorted(groupes_par_taille.keys(), reverse=True):
            if taille not in ordre_tailles_souhaite:
                ordre_tailles_souhaite.append(taille)

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
                                    
                                    options = get_options_filtrees(agres, row['Fonction'], agent_actuel)
                                    
                                    default_idx = 0
                                    for opt_idx, opt in enumerate(options):
                                        if agent_actuel.strip().lower() in opt.lower() and agent_actuel.strip() != "":
                                            default_idx = opt_idx
                                            break
                                    choix_label = st.selectbox(
                                        f"{agres}_{row['Fonction']}_{i}", 
                                        options=options, index=default_idx, 
                                        label_visibility="collapsed", key=f"agent_{i}_{agres}"
                                    )
                                    nouveau_personnel = dict_agents.get(choix_label, choix_label.split(" (")[0] if choix_label else "")
                                modifications_agents[(row['row_idx'], row['col_personnel'])] = nouveau_personnel
                            st.divider()

        with st.sidebar:
            st.markdown("### 📥 Actions")
            
            if os.path.exists("modele.xlsx"):
                try:
                    wb = openpyxl.load_workbook("modele.xlsx")
                    if 'BILLET' in wb.sheetnames:
                        ws = wb['BILLET']
                        
                        row_offset = 1
                        col_offset = 1
                        
                        if anchor_csv:
                            anchor_xls = None
                            for c in range(1, 30):
                                for r in range(1, 100):
                                    val = ws.cell(row=r, column=c).value
                                    if val and str(val).strip().upper() == anchor_csv[2]:
                                        anchor_xls = (r, c)
                                        break
                                if anchor_xls: break
                            
                            if anchor_xls:
                                row_offset = anchor_xls[0] - anchor_csv[0]
                                col_offset = anchor_xls[1] - anchor_csv[1]

                        for (r_csv, c_csv), val in modifications_agents.items():
                            ws.cell(row=r_csv + row_offset, column=c_csv + col_offset, value=val)
                    
                    output = BytesIO()
                    wb.save(output)
                    
                    st.download_button(
                        label="📥 Télécharger la Feuille Parfaite", 
                        data=output.getvalue(), 
                        file_name="Feuille_Garde_Finale.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"Erreur d'édition : {e}")
            else:
                st.error("❌ Fichier 'modele.xlsx' introuvable sur le serveur.")

except Exception as e:
    st.error(f"⚠️ Erreur Globale : {e}")
