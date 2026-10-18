import streamlit as st
import pandas as pd
import os
from datetime import datetime

# ==============================================================================
# 1. CONFIGURATION DE L'APPLICATION ET CONTEXTE SÉCURITÉ
# ==============================================================================
st.set_page_config(
    page_title="Gestion Flotte & Recettes", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- BASE DE DONNÉES CENTRALE DES COMPTES UTILISATEURS ---
COMPTES = {
    "dg": {"nom": "Directeur Général", "mp": "dg2026", "role": "DG"},
    "gerant": {"nom": "Gérant", "mp": "gerant2026", "role": "GERANT"},
    "adjoint": {"nom": "Adjoint", "mp": "adj2026", "role": "ADJOINT"},
    "secretaire": {"nom": "Secrétaire", "mp": "sec2026", "role": "SECRETAIRE"},
    "chauffeur": {"nom": "Chauffeur", "mp": "ch2026", "role": "CHAUFFEUR"}
}

# --- INITIALISATION RIGOUREUSE DE L'ÉTAT DE SESSION ---
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
if "user_role" not in st.session_state:
    st.session_state["user_role"] = None
if "user_name" not in st.session_state:
    st.session_state["user_name"] = None

# ==============================================================================
# 2. DIALOGUE D'AUTHENTIFICATION UNIQUE (ÉVITE L'ÉCRAN BLANC)
# ==============================================================================
if not st.session_state["authenticated"]:
    st.title("🔒 Accès Sécurisé")
    st.subheader("Veuillez entrer vos identifiants pour accéder à vos compartiments")
    
    with st.form(key="login_form_entreprise", clear_on_submit=True):
        username = st.text_input("Nom d'utilisateur").strip().lower()
        password = st.text_input("Mot de passe", type="password")
        submit_login = st.form_submit_button("Se connecter au réseau")
        
    if submit_login:
        if username in COMPTES and COMPTES[username]["mp"] == password:
            st.session_state["authenticated"] = True
            st.session_state["user_role"] = COMPTES[username]["role"]
            st.session_state["user_name"] = COMPTES[username]["nom"]
            st.success(f"✅ Bienvenue {st.session_state['user_name']} !")
            st.rerun()
        else:
            st.error("❌ Identifiant ou mot de passe incorrect.")
    st.stop()  # Arrêt strict uniquement ici tant que l'utilisateur n'est pas connecté

role_actif = st.session_state["user_role"]

# ==============================================================================
# 3. CONFLITS DE BASES DE DONNÉES ET SYNCHRONISATION EXCEL / CSV
# ==============================================================================
EXCEL_PARC = "gestion_parc_automobile.xlsx"
CSV_RECETTES = "recettes_officiel.csv"
CSV_TRAJETS = "trajets_officiel.csv"
CSV_FLOTTE = "flotte_officiel.csv"
CSV_CLIENTS = "clients_officiel.csv"
CSV_CHAUFFEURS = "chauffeurs_officiel.csv"
CSV_LOCATIONS = "locations_officiel.csv"

def synchroniser_et_charger(chemin_csv, colonnes_cibles, par_defaut=None):
    """Garantit le chargement transparent des données sans rupture de colonnes."""
    if os.path.exists(chemin_csv):
        try:
            df = pd.read_csv(chemin_csv)
            for c in colonnes_cibles:
                if c not in df.columns:
                    df[c] = None
            return df[colonnes_cibles]
        except Exception:
            pass
    if par_defaut is not None:
        return pd.DataFrame(par_defaut)
    return pd.DataFrame(columns=colonnes_cibles)

# Initialisation des structures en Session State
if "df_flotte" not in st.session_state:
    st.session_state["df_flotte"] = synchroniser_et_charger(CSV_FLOTTE, ["Date Enregistrement", "Immatriculation", "Marque / Modèle", "Type de Service", "Statut Véhicule"])

if "df_taxis" not in st.session_state:
    st.session_state["df_taxis"] = synchroniser_et_charger(CSV_RECETTES, ["Date", "Immatriculation", "Chauffeur", "Zone / Ligne", "Recette Brute (CFA)", "Carburant/Dépenses (CFA)", "Recette Nette (CFA)"])

if "df_clients" not in st.session_state:
    st.session_state["df_clients"] = synchroniser_et_charger(CSV_CLIENTS, ["Nom Complet", "Téléphone", "Commune"])

if "df_chauffeurs" not in st.session_state:
    st.session_state["df_chauffeurs"] = synchroniser_et_charger(CSV_CHAUFFEURS, ["ID Chauffeur", "Nom Complet", "Téléphone", "Numéro Permis", "Catégorie", "Expiration"])

if "df_loc" not in st.session_state:
    st.session_state["df_loc"] = synchroniser_et_charger(CSV_LOCATIONS, ["Code Location", "Date Début", "Date Fin", "Immatriculation", "Nom Complet", "Prix Total (CFA)"])

if "df_trajets" not in st.session_state:
    zones_init = [
        {"Zone / Ligne": "Zone 1: Gare Riviera 2 ↔ Gare Cocody St Jean"},
        {"Zone / Ligne": "Zone 2: Gare Palmeraie ↔ Gare Adjamé"},
        {"Zone / Ligne": "Zone 3: Gare Yopougon Maroc ↔ Gare Plateau"}
    ]
    st.session_state["df_trajets"] = synchroniser_et_charger(CSV_TRAJETS, ["Zone / Ligne"], zones_init)

# --- ANCRAGE CLASSEUR CENTRAL EXCEL AUTOMATIQUE ---
if not os.path.exists(EXCEL_PARC):
    try:
        with pd.ExcelWriter(EXCEL_PARC, engine="openpyxl") as master_writer:
            st.session_state["df_flotte"].to_excel(master_writer, sheet_name="Flotte", index=False)
            st.session_state["df_taxis"].to_excel(master_writer, sheet_name="Recettes_Taxis", index=False)
            st.session_state["df_loc"].to_excel(master_writer, sheet_name="Locations", index=False)
            st.session_state["df_clients"].to_excel(master_writer, sheet_name="Clients", index=False)
            st.session_state["df_chauffeurs"].to_excel(master_writer, sheet_name="Chauffeurs", index=False)
    except Exception:
        pass

def sauvegarder_donnees_globales():
    """Commit physique complet des structures vers le stockage disque local."""
    try:
        st.session_state["df_flotte"].to_csv(CSV_FLOTTE, index=False)
        st.session_state["df_taxis"].to_csv(CSV_RECETTES, index=False)
        st.session_state["df_clients"].to_csv(CSV_CLIENTS, index=False)
        st.session_state["df_chauffeurs"].to_csv(CSV_CHAUFFEURS, index=False)
        st.session_state["df_loc"].to_csv(CSV_LOCATIONS, index=False)
        st.session_state["df_trajets"].to_csv(CSV_TRAJETS, index=False)
        
        with pd.ExcelWriter(EXCEL_PARC, engine="openpyxl") as writer:
            st.session_state["df_flotte"].to_excel(writer, sheet_name="Flotte", index=False)
            st.session_state["df_taxis"].to_excel(writer, sheet_name="Recettes_Taxis", index=False)
            st.session_state["df_loc"].to_excel(writer, sheet_name="Locations", index=False)
            st.session_state["df_clients"].to_excel(writer, sheet_name="Clients", index=False)
            st.session_state["df_chauffeurs"].to_excel(writer, sheet_name="Chauffeurs", index=False)
        st.sidebar.success("💾 Synchronisation complète effectuée !")
    except Exception as error:
        st.sidebar.error(f"Erreur d'écriture disque : {error}")

# ==============================================================================
# 4. ARCHITECTURE NAVIGATION : BARRE LATÉRALE ET PRIVILÈGES
# ==============================================================================
with st.sidebar:
    st.write(f"👤 Connecté : **{st.session_state['user_name']}**")
    st.write(f"🔑 Rôle système : `{role_actif}`")
    st.markdown("---")
    
    options_menu = []
    
    if role_actif in ["DG", "GERANT"]:
        options_menu.append("📊 5. Calcul des Recettes Globales")
    if role_actif in ["DG", "GERANT", "ADJOINT", "SECRETAIRE", "CHAUFFEUR"]:
        options_menu.append("🚖 1. Recettes Taxis Communaux")
    if role_actif in ["DG", "GERANT", "ADJOINT", "SECRETAIRE"]:
        options_menu.append("🚗 2. Véhicules de Location & Prix")
        options_menu.append("👥 3. Identité des Clients")
    if role_actif in ["DG", "GERANT", "ADJOINT"]:
        options_menu.append("👨‍✈️ 4. Identité des Chauffeurs")
    if role_actif in ["DG", "GERANT"]:
        options_menu.append("⚙️ Configuration Flotte globale")
    if role_actif == "DG":
        options_menu.append("🔑 Gestion des Accès")

    st.markdown("### 🏬 Menu Principal")
    menu_selection = st.radio("📌 Navigation Principale", options_menu)
    st.markdown("---")
    
    if st.button("💾 Enregistrer les Modifications", use_container_width=True):
        sauvegarder_donnees_globales()
        
    if st.button("🚪 Se Déconnecter", use_container_width=True):
        st.session_state["authenticated"] = False
        st.rerun()

# ==============================================================================
# 5. INTEGRATION ET BRANCHEMENT DES COMPARTIMENTS OPERATIONNELS
# ==============================================================================
st.title(f"📍 {menu_selection}")

# Extraction dynamique des interdépendances pour les menus déroulants
flotte_active = st.session_state["df_flotte"]
liste_taxis = flotte_active[flotte_active["Type de Service"] == "Taxi Communal"]["Immatriculation"].tolist()
liste_locations = flotte_active[flotte_active["Type de Service"] == "Location"]["Immatriculation"].tolist()
liste_chauffeurs = st.session_state["df_chauffeurs"]["Nom Complet"].tolist()
liste_clients = st.session_state["df_clients"]["Nom Complet"].tolist()
liste_zones = st.session_state["df_trajets"]["Zone / Ligne"].tolist()

# --- DISPATCHING ROBUSTE DES COMPARTIMENTS ---

# ------------------------------------------------------------------------------
# 🚖 INTERFACE 1 : RECETTES TAXIS COMMUNAUX
# ------------------------------------------------------------------------------
if "1. Recettes Taxis Communaux" in menu_selection:
    st.header("Suivi quotidien opérationnel des Taxis Communaux")
    t_saisie, t_historique = st.tabs(["📥 Déclaration de Caisse", "📜 Historique"])
    
    with t_saisie:
        if not liste_taxis:
