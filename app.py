import streamlit as st
import pandas as pd
from io import BytesIO

# 1. Configuration de la page en mode large
st.set_page_config(page_title="Feuille de Garde - L'Isle-en-Dodon", page_icon="🚒", layout="wide")

# 2. Styles CSS personnalisés
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

    # Chargement de l'onglet des véhicules
    try:
        df_vehicules = pd.read_excel(fichier, sheet_name='VEHICULE')
    except Exception:
        df_vehicules = pd.DataFrame(columns=["ENGIN", "NUMERO"])

    return df_garde, df_vehicules

# Utilisation du nom de fichier propre sans espace
fichier_source = "feuille_garde.xlsx"

# En-tête de l'application
st.markdown('<p class="titre-caserne">🚒 Centre de Secours de L\'Isle-en-Dodon</p>', unsafe_allow_html=True)
st.markdown('<p class="sous-titre">Gestion opérationnelle des gardes</p>', unsafe_allow_html=True)

try:
    df_garde, df_vehicules = charger_donnees_excel(fichier_source)
    
    if df_garde.empty:
        st.warning("Le tableau de garde est vide. Vérifiez l'onglet 'BILLET'.")
    else:
        # --- TABLEAU PRINCIPAL EN PLEINE LARGEUR ---
        st.write("**📋 Affectation des équipages**")
        df_garde_modifie = st.data_editor(
            df_garde, 
            num_rows="dynamic", 
            use_container_width=True,
            height=600,
            hide_index=True 
        )

        # Nettoyage de la table véhicules pour les listes déroulantes
        df_veh_clean = df_vehicules.dropna(subset=['ENGIN']).copy()
        
        # Liste complète des numéros disponibles
        tous_les_numeros = [str(int(n)) for n in df_vehicules['NUMERO'].dropna() if str(n).replace('.','',1).isdigit()]
        if not tous_les_numeros:
            tous_les_numeros = [str(i) for i in range(1, 100)]

        # --- BARRE LATÉRALE (SIDEBAR) AVEC LISTES DÉROULANTES ---
        effectif_saisi = df_garde_modifie['Personnel'].apply(lambda x: 1 if str(x).strip() != "" and str(x).lower() != "nan" else 0).sum()
        
        with st.sidebar:
            st.header("⚙️ Actions")
            st.metric(label="Postes pourvus", value=f"{effectif_saisi} / {len(df_garde)}")
            st.divider()
            
            with st.expander("🚛 Assigner les numéros de véhicules"):
                nouveaux_vehicules = []
                for idx, row in df_veh_clean.iterrows():
                    engin_nom = row['ENGIN']
                    ancien_num = str(int(row['NUMERO'])) if pd.notna(row['NUMERO']) and str(row['NUMERO']).replace('.','',1).isdigit() else tous_les_numeros[0]
                    
                    if ancien_num not in tous_les_numeros:
                        tous_les_numeros.insert(0, ancien_num)
                        
                    index_defaut = tous_les_numeros.index(ancien_num) if ancien_num in tous_les_numeros else 0
                    
                    choix_num = st.selectbox(f"Indicatif {engin_nom}", tous_les_numeros, index=index_defaut, key=f"veh_{engin_nom}_{idx}")
                    nouveaux_vehicules.append({"ENGIN": engin_nom, "NUMERO": choix_num})
                
                df_vehicules_modifie = pd.DataFrame(nouveaux_vehicules)

            st.divider()
            
            # --- PRÉPARATION DU FICHIER EXCEL DE SORTIE ---
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_garde_modifie.to_excel(writer, index=False, sheet_name='Garde')
                df_vehicules_modifie.to_excel(writer, index=False, sheet_name='VEHICULE')
            
            st.download_button(
                label="📥 Télécharger la feuille validée", 
                data=output.getvalue(), 
                file_name="Feuille_Garde_Modifiee.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary" 
            )

except Exception as e:
    st.error(f"⚠️ Erreur : {e}")
