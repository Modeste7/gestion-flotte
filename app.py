import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta

# ==============================================================================
# 1. INITIALISATION DE L'INTERFACE ET CONFIGURATION DE PAGE
# ==============================================================================
st.set_page_config(
    page_title="Gestion Flotte & Recettes Pro", 
    layout="wide", 
    initial_sidebar_state="expanded"
)

# --- BASE DE DONNÉES DES UTILISATEURS ET DES RÔLES ACCÈS ---
COMPTES = {
    "dg": {"nom": "Directeur Général", "mp": "dg2026", "role": "DG"},
    "gerant": {"nom": "Gérant", "mp": "gerant2026", "role": "GERANT"},
    "adjoint": {"nom": "Adjoint", "mp": "adj2026", "role": "ADJOINT"},
    "secretaire": {"nom": "Secrétaire", "mp": "sec2026", "role": "SECRETAIRE"},
    "chauffeur": {"nom": "Chauffeur", "mp": "ch2026", "role": "CHAUFFEUR"}
}

# --- INITIALISATION SÉCURISÉE DES VARIABLES DE SESSION ---
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
if "user_role" not in st.session_state:
    st.session_state["user_role"] = None
if "user_name" not in st.session_state:
    st.session_state["user_name"] = None

# --- CONSTANTES DES CHEMINS DE STOCKAGE (CSV & EXCEL) ---
EXCEL_PARC = "gestion_parc_automobile.xlsx"
CSV_FILE_RECETTES = "recettes_officiel.csv"
CSV_FILE_TRAJETS = "trajets_officiel.csv"
CSV_FILE_FLOTTE = "flotte_officiel.csv"
CSV_FILE_CLIENTS = "clients_officiel.csv"
CSV_FILE_CHAUFFEURS = "chauffeurs_officiel.csv"
CSV_FILE_LOCATIONS = "locations_officiel.csv"

# ==============================================================================
# 2. ÉCRAN DE VÉRIFICATION ET D'AUTHENTIFICATION UNIQUE
# ==============================================================================
if not st.session_state["authenticated"]:
    st.title("🔒 Accès Sécurisé au Système")
    st.subheader("Veuillez entrer vos identifiants d'entreprise pour déverrouiller vos modules")
    
    with st.form(key="formulaire_connexion_principal", clear_on_submit=True):
        username_saisi = st.text_input("Nom d'utilisateur").strip().lower()
        password_saisi = st.text_input("Mot de passe", type="password")
        bouton_connexion = st.form_submit_button("Se connecter au réseau")
        
    if bouton_connexion:
        if username_saisi in COMPTES and COMPTES[username_saisi]["mp"] == password_saisi:
            st.session_state["authenticated"] = True
            st.session_state["user_role"] = COMPTES[username_saisi]["role"]
            st.session_state["user_name"] = COMPTES[username_saisi]["nom"]
            st.success(f"✅ Connexion établie. Bienvenue {st.session_state['user_name']} !")
            st.rerun()
        else:
            st.error("❌ Identifiant ou mot de passe incorrect.")
    st.stop()

# Assignation du rôle pour le reste du script
role_actif = st.session_state["user_role"]

# ==============================================================================
# 3. MOTEUR DE CHARGEMENT DE DONNÉES ET GESTION DES FICHIERS
# ==============================================================================
def charger_base_donnees(chemin_csv, liste_colonnes, donnees_defaut=None):
    """Charge un fichier CSV s'il existe, sinon l'initialise proprement en mémoire."""
    if os.path.exists(chemin_csv):
        try:
            donnees_chargees = pd.read_csv(chemin_csv)
            for colonne in liste_colonnes:
                if colonne not in donnees_chargees.columns:
                    donnees_chargees[colonne] = None
            return donnees_chargees[liste_colonnes]
        except Exception:
            pass
    if donnees_defaut is not None:
        return pd.DataFrame(donnees_defaut)
    return pd.DataFrame(columns=liste_colonnes)

# Chargement en mémoire vive (Session State)
if "df_flotte" not in st.session_state:
    st.session_state["df_flotte"] = charger_base_donnees(CSV_FILE_FLOTTE, ["Date Enregistrement", "Immatriculation", "Marque / Modèle", "Type de Service", "Statut Véhicule"])

if "df_taxis" not in st.session_state:
    st.session_state["df_taxis"] = charger_base_donnees(CSV_FILE_RECETTES, ["Date", "Immatriculation", "Chauffeur", "Zone / Ligne", "Recette Brute (CFA)", "Carburant/Dépenses (CFA)", "Recette Nette (CFA)"])

if "df_clients" not in st.session_state:
    st.session_state["df_clients"] = charger_base_donnees(CSV_FILE_CLIENTS, ["ID Client", "Nom Complet", "Téléphone", "Adresse"])

if "df_chauffeurs" not in st.session_state:
    st.session_state["df_chauffeurs"] = charger_base_donnees(CSV_FILE_CHAUFFEURS, ["ID Chauffeur", "Nom Complet", "Téléphone", "Numéro Permis", "Catégorie", "Expiration"])

if "df_loc" not in st.session_state:
    st.session_state["df_loc"] = charger_base_donnees(CSV_FILE_LOCATIONS, ["Code Location", "Date Début", "Date Fin", "Immatriculation", "Nom Complet", "Prix Total (CFA)"])

if "df_trajets" not in st.session_state:
    lignes_predefinies = [
        {"Zone / Ligne": "Zone 1: Gare Riviera 2 ↔ Gare Cocody St Jean"},
        {"Zone / Ligne": "Zone 2: Gare Palmeraie ↔ Gare Adjamé"},
        {"Zone / Ligne": "Zone 3: Gare Yopougon Maroc ↔ Gare Plateau"}
    ]
    st.session_state["df_trajets"] = charger_base_donnees(CSV_FILE_TRAJETS, ["Zone / Ligne"], lignes_predefinies)

# --- SÉCURITÉ : INITIALISATION DU MASTER FILE EXCEL SI ABSENT ---
if not os.path.exists(EXCEL_PARC):
    try:
        with pd.ExcelWriter(EXCEL_PARC, engine="openpyxl") as ecrivain_excel:
            st.session_state["df_flotte"].to_excel(ecrivain_excel, sheet_name="Flotte", index=False)
            st.session_state["df_taxis"].to_excel(ecrivain_excel, sheet_name="Recettes_Taxis", index=False)
            st.session_state["df_loc"].to_excel(ecrivain_excel, sheet_name="Locations", index=False)
            st.session_state["df_clients"].to_excel(ecrivain_excel, sheet_name="Clients", index=False)
            st.session_state["df_chauffeurs"].to_excel(ecrivain_excel, sheet_name="Chauffeurs", index=False)
    except Exception as erreur_initialisation:
        st.error(f"Erreur d'initialisation du fichier Excel maître : {erreur_initialisation}")

# ==============================================================================
# 4. FONCTION DE PERSISTANCE ATOMIQUE (CRITURE CONCURRENTE)
# ==============================================================================
def enregistrer_modifications_locales():
    """Sauvegarde de manière synchrone toutes les bases de données mémoires vers CSV et Excel."""
    try:
        # Étape A: Écriture des bases au format CSV individuel
        st.session_state["df_flotte"].to_csv(CSV_FILE_FLOTTE, index=False)
        st.session_state["df_taxis"].to_csv(CSV_FILE_RECETTES, index=False)
        st.session_state["df_clients"].to_csv(CSV_FILE_CLIENTS, index=False)
        st.session_state["df_chauffeurs"].to_csv(CSV_FILE_CHAUFFEURS, index=False)
        st.session_state["df_loc"].to_csv(CSV_FILE_LOCATIONS, index=False)
        st.session_state["df_trajets"].to_csv(CSV_FILE_TRAJETS, index=False)
        
        # Étape B: Compilation générale dans le fichier Excel d'entreprise
        with pd.ExcelWriter(EXCEL_PARC, engine="openpyxl") as synchro_excel:
            st.session_state["df_flotte"].to_excel(synchro_excel, sheet_name="Flotte", index=False)
            st.session_state["df_taxis"].to_excel(synchro_excel, sheet_name="Recettes_Taxis", index=False)
            st.session_state["df_loc"].to_excel(synchro_excel, sheet_name="Locations", index=False)
            st.session_state["df_clients"].to_excel(synchro_excel, sheet_name="Clients", index=False)
            st.session_state["df_chauffeurs"].to_excel(synchro_excel, sheet_name="Chauffeurs", index=False)
            st.session_state["df_trajets"].to_excel(synchro_excel, sheet_name="Trajets_Réseaux", index=False)
            
        st.sidebar.success("💾 Base de données synchronisée (CSV & Excel) !")
    except Exception as e:
        st.sidebar.error(f"❌ Échec de la sauvegarde sur le disque : {e}")

# ==============================================================================
# 5. CONTRÔLE DE NAVIGATION LATÉRALE ET COMPARTIMENTAGE DYNAMIQUE
# ==============================================================================
with st.sidebar:
    st.markdown(f"### 👤 {st.session_state['user_name']}")
    st.markdown(f"Rôle d'accès : `{role_actif}`")
    st.markdown("---")
    
    # Filtrage algorithmique des droits d'affichage selon le rôle utilisateur
    routes_autorisees = []
    
    if role_actif in ["DG", "GERANT"]:
        routes_autorisees.append("📊 5. Calcul des Recettes Globales")
        
    if role_actif in ["DG", "GERANT", "ADJOINT", "SECRETAIRE", "CHAUFFEUR"]:
        routes_autorisees.append("🚖 1. Recettes Taxis Communaux")
        
    if role_actif in ["DG", "GERANT", "ADJOINT", "SECRETAIRE"]:
        routes_autorisees.append("🚗 2. Véhicules de Location & Prix")
        routes_autorisees.append("👥 3. Identité des Clients")
        
    if role_actif in ["DG", "GERANT", "ADJOINT"]:
        routes_autorisees.append("👨‍✈️ 4. Identité des Chauffeurs")
        
    if role_actif in ["DG", "GERANT"]:
        routes_autorisees.append("⚙️ Configuration Flotte globale")
        
    if role_actif == "DG":
        routes_autorisees.append("🔑 Gestion des Accès")

    st.markdown("### 🏬 Navigation Système")
    choix_menu = st.radio("Sélectionnez votre module opérationnel :", routes_autorisees)
    st.markdown("---")
    
    # Actions rapides de la barre de contrôle latérale
    if st.button("💾 Enregistrer toutes les écritures", use_container_width=True):
        enregistrer_modifications_locales()
        
    if st.button("🚪 Fermer la Session", use_container_width=True):
        st.session_state["authenticated"] = False
        st.session_state["user_role"] = None
        st.session_state["user_name"] = None
        st.rerun()

# ==============================================================================
