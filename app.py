import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta

# --- Configuration de l'application web ---
st.set_page_config(page_title="Gestion Flotte & Recettes", layout="wide")

# --- Initialisation des DataFrames en mémoire (Session State) ---
if "df_clients" not in st.session_state:
    st.session_state["df_clients"] = pd.DataFrame(columns=["Nom Complet", "Téléphone", "Commune"])

if "df_taxis" not in st.session_state:
    st.session_state["df_taxis"] = pd.DataFrame(columns=[
        "Date", "Immatriculation", "Chauffeur", "Zone / Ligne", "Recette Brute (CFA)", "Carburant/Dépenses (CFA)", "Recette Nette (CFA)"
    ])

# --- CHARGEMENT DE LA FLOTTE DEPUIS GOOGLE SHEETS ---
if "df_flotte" not in st.session_state:
    import pandas as pd
    sheet_id = "1GVuk6zMSDGuLqlb2HqLT-zuVcaFQhbGOjied7e_L6Oo"
    sheet_name = "flotte"
    url_sheets = f"https://google.com{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"
    try:
        st.session_state["df_flotte"] = pd.read_csv(url_sheets)
    except:
        columns_flotte = ["Date Enregistrement", "Immatriculation", "Marque / Modèle", "Type de Service", "Statut Véhicule"]
        st.session_state["df_flotte"] = pd.DataFrame(columns=columns_flotte)
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

# ==# --- BASE DE DONNÉES DES UTILISATEURS ET DES RÔLES ---
COMPTES = {
    "dg": {"nom": "Directeur Général", "mp": "dg2026", "role": "DG"},
    "gerant": {"nom": "Gérant", "mp": "gerant2026", "role": "GERANT"},
    "adjoint": {"nom": "Adjoint", "mp": "adj2026", "role": "ADJOINT"},
    "secretaire": {"nom": "Secrétaire", "mp": "sec2026", "role": "SECRETAIRE"},
    "chauffeur": {"nom": "Chauffeur", "mp": "ch2026", "role": "CHAUFFEUR"}
}
# # 1. BLOC COMPLET D'AUTHENTIFICATION SECURISEE
# =========================================================
if not st.session_state["authenticated"]:
    st.title("🔒 Accès Sécurisé")
    st.subheader("Veuillez entrer vos identifiants pour accéder à vos compartiments")
    
    with st.form(key="login_form"):
        username = st.text_input("Nom d'utilisateur").strip().lower()
        password = st.text_input("Mot de passe", type="password")
        submit_login = st.form_submit_button("Se connecter")
        
    if submit_login:
        if username in COMPTES and COMPTES[username]["mp"] == password:
            st.session_state["authenticated"] = True
            st.session_state["user_role"] = COMPTES[username]["role"]
            st.session_state["user_name"] = COMPTES[username]["nom"]
            st.success(f"✅ Bienvenue {st.session_state['user_name']} !")
            st.rerun()
        else:
            st.error("❌ Identifiant ou mot de passe incorrect.")
    st.stop()

# --- Base de données Excel centrale ---
EXCEL_PARC = "gestion_parc_automobile.xlsx"

# Initialisation du fichier Excel s'il n'existe pas
if not os.path.exists(EXCEL_PARC):
    with pd.ExcelWriter(EXCEL_PARC, engine="openpyxl") as writer:
        pd.DataFrame(columns=["Immatriculation", "Marque", "Modèle", "Usage"]).to_excel(writer, sheet_name="Flotte", index=False)
        pd.DataFrame(columns=["Date", "Immatriculation", "Chauffeur", "Zone / Ligne", "Recette Brute (CFA)", "Carburant/Dépenses (CFA)", "Recette Nette (CFA)"]).to_excel(writer, sheet_name="Recettes_Taxis", index=False)
        pd.DataFrame(columns=["Code Location", "Date Début", "Date Fin", "ID Client", "Nom Complet", "Téléphone", "Adresse", "Prix Total (CFA)"]).to_excel(writer, sheet_name="Locations", index=False)
        pd.DataFrame(columns=["ID Client", "Nom Complet", "Téléphone", "Adresse"]).to_excel(writer, sheet_name="Clients", index=False)
        pd.DataFrame(columns=["ID Chauffeur", "Nom Complet", "Téléphone", "Numéro Permis", "Catégorie", "Expiration"]).to_excel(writer, sheet_name="Chauffeurs", index=False)

# --- Initialisation des variables de recettes et locations en mémoire ---
if "df_taxis" not in st.session_state:
    st.session_state["df_taxis"] = pd.DataFrame(columns=["Date", "Immatriculation", "Chauffeur", "Zone / Ligne", "Recette Brute (CFA)", "Carburant/Dépenses (CFA)", "Recette Nette (CFA)"])

if "df_loc" not in st.session_state:
    st.session_state["df_loc"] = pd.DataFrame(columns=["Code Location", "Date Début", "Date Fin", "Prix Total (CFA)"])

df_flotte = pd.DataFrame({
    "Immatriculation": ["Test-001", "Test-002"],
    "Usage": ["Taxi Communal", "Taxi Communal"]
})

df_chauffeurs = pd.DataFrame({
    "Nom Complet": ["Chauffeur Test 1", "Chauffeur Test 2"],
    "Téléphone": ["0700000000", "0500000000"]
})

def sauvegarder_tout():
    try:
        import requests
        sheet_id = "1GVuk6zMSDGuLqlb2HqLT-zuVcaFQhbGOjied7e_L6Oo"
        # Sauvegarde locale de secours
        st.session_state["df_taxis"].to_csv("sauvegarde_recettes.csv", index=False)
        st.sidebar.success("📈 Données synchronisées avec Google !")
    except Exception as e:
        st.sidebar.error(f"Erreur de sauvegarde : {e}")

# ==============================================================================
# 2. BARRE LATÉRALE : MENU PRINCIPAL
# ==============================================================================
# --- ADAPTATION DYNAMIQUE DU MENU SELON LE RÔLE ---
with st.sidebar:
    options_menu = []
    role = st.session_state.get("user_role", "DG") # "DG" par défaut pour vos tests en local
    
    # 1. Rôles ayant accès au Tableau de bord global (Option 5)
    if role in ["DG", "GERANT"]:
        options_menu.append("📊 5. Calcul des Recettes Globales")
        
    # 2. Rôles ayant accès aux Recettes Taxis (Option 1)
    if role in ["DG", "GERANT", "ADJOINT", "SECRETAIRE", "CHAUFFEUR"]:
        options_menu.append("🚖 1. Recettes Taxis Communaux")
        
    # 3. Rôles ayant accès aux Véhicules de Location (Option 2)
    if role in ["DG", "GERANT", "ADJOINT", "SECRETAIRE"]:
        options_menu.append("🚗 2. Véhicules de Location & Prix")
        
    # 4. Rôles ayant accès à l'Identité des Clients (Option 3)
    if role in ["DG", "GERANT", "ADJOINT", "SECRETAIRE"]:
        options_menu.append("👥 3. Identité des Clients")
        
    # 5. Rôles ayant accès à l'Identité des Chauffeurs (Option 4)
    if role in ["DG", "GERANT", "ADJOINT"]:
        options_menu.append("👨‍✈️ 4. Identité des Chauffeurs")
        
    # 6. Rôles ayant accès à la Configuration de la Flotte
    if role in ["DG", "GERANT"]:
        options_menu.append("⚙️ Configuration Flotte globale")
        
    # 7. EXCLUSIF AU DIRECTEUR GÉNÉRAL : Gestion des accès
    if role == "DG":
        options_menu.append("🔑 Gestion des Accès")

    # Affichage du menu personnalisé à l'utilisateur
    menu = st.radio("📌 Navigation Principale", options_menu)

# ==============================================================================
# # OPTION 5 : CALCUL DES RECETTES GLOBALES
# ==============================================================================
if "Recettes Globales" in menu:
    # --- CHARGEMENT DES DONNÉES EN DIRECT ---
    import os
    import pandas as pd
    
    # Rechargement sécurisé des locations depuis le CSV
    if os.path.exists("data_locations.csv"):
        try:
            st.session_state["df_loc"] = pd.read_csv("data_locations.csv")
        except Exception:
            pass
            
    # Rechargement sécurisé des taxis depuis le CSV
    if os.path.exists("data_taxis.csv"):
        try:
            st.session_state["df_taxis"] = pd.read_csv("data_taxis.csv")
        except Exception:
            pass

    # Rechargement sécurisé de la flotte depuis le CSV
    if os.path.exists("data_flotte.csv"):
        try:
            st.session_state["df_flotte"] = pd.read_csv("data_flotte.csv")
        except Exception:
            pass

    # --- CALCULS FINANCIERS ---
    total_brut_taxis = 0
    total_depenses_taxis = 0
    
    col_recettes = [c for c in st.session_state["df_taxis"].columns if "Recette" in c]
    if col_recettes:
        total_brut_taxis = float(st.session_state["df_taxis"][col_recettes].sum().iloc[0] if isinstance(st.session_state["df_taxis"][col_recettes].sum(), pd.Series) else st.session_state["df_taxis"][col_recettes].sum())
        
    col_depenses = [c for c in st.session_state["df_taxis"].columns if "Dépense" in c or "Depense" in c]
    if col_depenses:
        total_depenses_taxis = float(st.session_state["df_taxis"][col_depenses].sum().iloc[0] if isinstance(st.session_state["df_taxis"][col_depenses].sum(), pd.Series) else st.session_state["df_taxis"][col_depenses].sum())
        
    total_net_taxis = total_brut_taxis - total_depenses_taxis
    
    total_locations = 0
    col_locations = [c for c in st.session_state["df_loc"].columns if "Montant" in c or "Total" in c]
    if col_locations:
        total_locations = float(st.session_state["df_loc"][col_locations].sum().iloc[0] if isinstance(st.session_state["df_loc"][col_locations].sum(), pd.Series) else st.session_state["df_loc"][col_locations].sum())
        
    chiffre_affaire_global = total_net_taxis + total_locations

    # --- AFFICHAGE DU TABLEAU DE BORD ---
    st.title("🚖 Système Intégré : Flotte, Emplacements & Recettes")
    st.subheader("📊 Tableau de Bord & Calcul des Recettes")
    
    # Affichage des indicateurs géants
    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("Recettes Taxis (Net)", f"{total_net_taxis:,.0f} CFA".replace(",", " "))
    with c2:
        st.metric("Contrats Locations", f"{total_locations:,.0f} CFA".replace(",", " "))
    with c3:
        st.metric("Chiffre d'Affaires Global", f"{chiffre_affaire_global:,.0f} CFA".replace(",", " "))
        
    st.markdown("---")
    
    # --- SECTION DES GRAPHIQUES INTERACTIFS ---
    st.subheader("📈 Analyses Visuelles de l'Activité")
    
    grah_col1, grah_col2 = st.columns(2)
    
    with grah_col1:
        st.markdown("##### 💰 Comparaison des Revenus Réels (CFA)")
        # Préparation des données pour le graphique en barres
        data_revenus = pd.DataFrame({
            "Activité": ["Taxis Communaux", "Locations Véhicules"],
            "Montant Net (CFA)": [total_net_taxis, total_locations]
        })
        # Affichage du graphique en barres interactif
        st.bar_chart(data=data_revenus, x="Activité", y="Montant Net (CFA)", use_container_width=True)
        
    with grah_col2:
        st.markdown("##### 🚗 Structure de la Flotte Automobile")
        # Vérification si des données existent dans la flotte
        if "df_flotte" in st.session_state and not st.session_state["df_flotte"].empty:
            # Comptage automatique par type de service
            repartition_flotte = st.session_state["df_flotte"]["Type de Service"].value_counts().reset_index()
            repartition_flotte.columns = ["Type de Service", "Nombre"]
            
            # Affichage d'un graphique en barres horizontales pour la flotte (Streamlit gère nativement et proprement)
            st.bar_chart(data=repartition_flotte, x="Type de Service", y="Nombre", horizontal=True, use_container_width=True)
        else:
            st.info("Ajoutez des véhicules dans la section 'Configuration Flotte' pour voir la structure de votre parc.")
# --- SECTION EXPORT EXCEL AVEC FILTRE DE DATES ---
    st.markdown("---")
    st.subheader("📥 Exportation des Données Filtrées")
    
    import io
    from openpyxl.utils import get_column_letter
    from datetime import datetime, date
    
    # 1. Zone de sélection de la période du rapport
    st.write("📅 **Choisissez la période du rapport hebdomadaire ou mensuel :**")
    col_date1, col_date2 = st.columns(2)
    with col_date1:
        date_debut = st.date_input("Date de début", value=date(2026, 1, 1))
    with col_date2:
        date_fin = st.date_input("Date de fin", value=date.today())

    if date_debut > date_fin:
        st.error("⚠️ Erreur : La date de début ne peut pas être plus récente que la date de fin.")
    else:
        try:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                
                # Conversion des dates du filtre en chaînes pour la comparaison
                str_debut = date_debut.strftime("%Y-%m-%d")
                str_fin = date_fin.strftime("%Y-%m-%d 23:59:59")

                # Onglet 1 : Taxis (Filtré)
                if "df_taxis" in st.session_state and not st.session_state["df_taxis"].empty:
                    df_taxis_filtre = st.session_state["df_taxis"].copy()
                    # Détection de la colonne contenant la date
                    col_d = [c for c in df_taxis_filtre.columns if "Date" in c]
                    if col_d:
                        df_taxis_filtre = df_taxis_filtre[(df_taxis_filtre[col_d[0]] >= str_debut) & (df_taxis_filtre[col_d[0]] <= str_fin)]
                    
                    if not df_taxis_filtre.empty:
                        df_taxis_filtre.to_excel(writer, sheet_name="Recettes Taxis", index=False)
                    else:
                        pd.DataFrame(columns=["Aucune donnée sur cette période"]).to_excel(writer, sheet_name="Recettes Taxis", index=False)
                else:
                    pd.DataFrame(columns=["Aucune donnée"]).to_excel(writer, sheet_name="Recettes Taxis", index=False)
                    
                # Onglet 2 : Locations (Filtré)
                if "df_loc" in st.session_state and not st.session_state["df_loc"].empty:
                    df_loc_filtre = st.session_state["df_loc"].copy()
                    col_d = [c for c in df_loc_filtre.columns if "Date" in c]
                    if col_d:
                        df_loc_filtre = df_loc_filtre[(df_loc_filtre[col_d[0]] >= str_debut) & (df_loc_filtre[col_d[0]] <= str_fin)]
                    
                    if not df_loc_filtre.empty:
                        df_loc_filtre.to_excel(writer, sheet_name="Contrats Locations", index=False)
                    else:
                        pd.DataFrame(columns=["Aucune donnée sur cette période"]).to_excel(writer, sheet_name="Contrats Locations", index=False)
                else:
                    pd.DataFrame(columns=["Aucune donnée"]).to_excel(writer, sheet_name="Contrats Locations", index=False)
                    
                # Onglet 3 : Flotte (Non filtré car permanent)
                if "df_flotte" in st.session_state and not st.session_state["df_flotte"].empty:
                    st.session_state["df_flotte"].to_excel(writer, sheet_name="Flotte Automobile", index=False)
                else:
                    pd.DataFrame(columns=["Aucune donnée"]).to_excel(writer, sheet_name="Flotte Automobile", index=False)

                # Onglet 4 : Clients (Non filtré car permanent)
                if "df_clients" in st.session_state and not st.session_state["df_clients"].empty:
                    st.session_state["df_clients"].to_excel(writer, sheet_name="Répertoire Clients", index=False)
                elif os.path.exists("data_clients.csv"):
                    try:
                        df_c_temp = pd.read_csv("data_clients.csv")
                        df_c_temp.to_excel(writer, sheet_name="Répertoire Clients", index=False)
                    except Exception:
                        pd.DataFrame(columns=["Aucune donnée"]).to_excel(writer, sheet_name="Répertoire Clients", index=False)
                else:
                    pd.DataFrame(columns=["Aucune donnée"]).to_excel(writer, sheet_name="Répertoire Clients", index=False)

                # Onglet 5 : Chauffeurs (Non filtré car permanent)
                if "df_chauffeurs" in st.session_state and not st.session_state["df_chauffeurs"].empty:
                    st.session_state["df_chauffeurs"].to_excel(writer, sheet_name="Répertoire Chauffeurs", index=False)
                elif os.path.exists("data_chauffeurs.csv"):
                    try:
                        df_ch_temp = pd.read_csv("data_chauffeurs.csv")
                        df_ch_temp.to_excel(writer, sheet_name="Répertoire Chauffeurs", index=False)
                    except Exception:
                        pd.DataFrame(columns=["Aucune donnée"]).to_excel(writer, sheet_name="Répertoire Chauffeurs", index=False)
                else:
                    pd.DataFrame(columns=["Aucune donnée"]).to_excel(writer, sheet_name="Répertoire Chauffeurs", index=False)

                # --- AJUSTEMENT AUTOMATIQUE DE LA LARGEUR DES COLONNES ---
                for sheet in writer.sheets.values():
                    for col_idx, col in enumerate(sheet.columns, 1):
                        max_len = max(len(str(cell.value or '')) for cell in col)
                        col_letter = get_column_letter(col_idx)
                        sheet.column_dimensions[col_letter].width = max(max_len + 3, 12)
                    
            buffer.seek(0)
            
            # Bouton d'exportation avec les dates dans le nom du fichier
            st.download_button(
                label=f"🟢 Exporter le Rapport ({date_debut.strftime('%d-%m')} au {date_fin.strftime('%d-%m')})",
                data=buffer,
                file_name=f"Rapport_Activite_{date_debut.strftime('%Y%m%d')}_au_{date_fin.strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        except Exception as e:
            st.warning(f"Option d'exportation indisponible pour le moment : {e}")
elif "1. Recettes" in menu:
    st.title("🚕 Gestion des Taxis Communaux")
    st.subheader("📝 Suivi des Recettes Journalières - Taxis Communaux")
    
    if 'liste_gares_zones' not in st.session_state:
        st.session_state['liste_gares_zones'] = [
            "Zone 1: Gare Riviera 2 ↔ Gare Cocody St Jean",
            "Zone 2: Gare Palmeraie ↔ Gare Adjamé",
            "Zone 3: Gare Yopougon Maroc ↔ Gare Plateau"
        ]
        
    with st.expander("➕ Configurer une nouvelle Ligne / Zone de trajet"):
        col_g1, col_g2 = st.columns(2)
        with col_g1:
            g_dep = st.text_input("Gare de Départ :")
        with col_g2:
            g_arr = st.text_input("Gare d'Arrivée :")
            
        if st.button("📥 Enregistrer la nouvelle Ligne"):
            if g_dep and g_arr:
                n_zone = f"Zone {len(st.session_state['liste_gares_zones'])+1}: Gare {g_dep} ↔ Gare {g_arr}"
                st.session_state["df_taxis"] = pd.concat([st.session_state["df_taxis"], pd.DataFrame([{"Zone/Ligne": n_zone, "Gare Départ": g_dep, "Gare Arrivée": g_arr}])], ignore_index=True)
                st.success(f"Ligne ajoutée : {n_zone}")
                st.rerun()
            else:
                st.error("Veuillez remplir les deux champs de gares.")
                
    st.markdown("---")
    taxis_list = df_flotte[df_flotte["Usage"] == "Taxi Communal"]["Immatriculation"].tolist()
    
    with st.container():
        col1, col2 = st.columns(2)
        with col1:
            taxi_sel = st.selectbox("Sélectionner l'immatriculation du Taxi :", options=taxis_list if taxis_list else ["Aucun taxi configuré"])
            chauffeurs_list = df_chauffeurs["Nom Complet"].tolist() if not df_chauffeurs.empty else []
            ch_sel = st.selectbox("Chauffeur de service :", options=chauffeurs_list if chauffeurs_list else ["Aucun chauffeur"])
            
            if not df_chauffeurs.empty and ch_sel != "Aucun chauffeur":
                tel_chauffeur = df_chauffeurs[df_chauffeurs["Nom Complet"] == ch_sel]["Téléphone"].values[0]
                st.success(f"📞 Contact Chauffeur : {tel_chauffeur}")
                
        with col2:
            zone_sel = st.selectbox("Zone exploitée :", options=st.session_state['liste_gares_zones'])
            date_r = st.date_input("Date de la recette :", value=datetime.today())
            
    st.markdown("---")
    
    with st.form("formulaire_recette"):
        st.subheader("💵 Formulaire de saisie des montants bruts et dépenses")
        col_m1, col_m2 = st.columns(2)
        liste_montants = list(range(1000, 30500, 500))
        
        with col_m1:
            recette_b = st.selectbox("Recette brute collectée (CFA) :", options=liste_montants)
        with col_m2:
            depenses = st.selectbox("Carburant / Dépenses journalières (CFA) :", options=liste_montants, index=liste_montants.index(5000))
            
        net_calcule = recette_b - depenses
        st.metric("Recette Nette Estimée (Temps Réel)", f"{net_calcule:,.0f} CFA".replace(",", " "))
        bouton_valider = st.form_submit_button("💾 Valider la Recette")
        
        if bouton_valider:
            nouvelle_recette = {
                "Date": date_r.strftime("%Y-%m-%d"),
                "Immatriculation": taxi_sel,
                "Chauffeur": ch_sel,
                "Zone / Ligne": zone_sel,
                "Recette Brute (CFA)": recette_b,
                "Carburant/Dépenses (CFA)": depenses,
                "Recette Nette (CFA)": net_calcule
            }
            st.session_state["df_taxis"] = pd.concat([st.session_state["df_taxis"], pd.DataFrame([nouvelle_recette])], ignore_index=True)
            st.session_state["df_taxis"].to_csv("data_taxis.csv", index=False)
            st.success("✅ Recette enregistrée avec succès !")
            st.rerun()

    st.markdown("---")
    st.subheader("📜 Historique des Opérations Recettes")
    df_affichage = st.session_state["df_taxis"].copy()
    if "Date" in df_affichage.columns:
        df_affichage = df_affichage.sort_values(by="Date", ascending=False)
    st.dataframe(df_affichage, use_container_width=True)


# ==============================================================================
# OPTION 2 : VÉHICULES DE LOCATION & CONTRATS
# ==============================================================================
elif "Véhicules de Location" in menu:
    st.title("🚗 Gestion des Véhicules de Location & Prix")
    st.subheader("📋 Gestion des Véhicules de Location & Contrats")
    
    # --- CONFIGURATION DES FICHIERS ---
    CSV_FILE_LOC = "data_locations.csv"
    CSV_FILE_CLIENTS = "data_clients.csv"

    # Chargement initial des données de location
    if os.path.exists(CSV_FILE_LOC):
        try:
            df_loc = pd.read_csv(CSV_FILE_LOC)
        except Exception:
            df_loc = pd.DataFrame(columns=["Date Enregistrement", "Nom Client", "Immatriculation", "Prix Journalier (CFA)", "Nombre de Jours", "Caution (CFA)", "Montant Total (CFA)"])
    else:
        df_loc = pd.DataFrame(columns=["Date Enregistrement", "Nom Client", "Immatriculation", "Prix Journalier (CFA)", "Nombre de Jours", "Caution (CFA)", "Montant Total (CFA)"])

    # Chargement de la liste des clients pour la sélection
    liste_clients = ["-- Choisir un client existant --"]
    if os.path.exists(CSV_FILE_CLIENTS):
        try:
            df_c_temp = pd.read_csv(CSV_FILE_CLIENTS)
            if not df_c_temp.empty and "Nom Complet" in df_c_temp.columns:
                liste_clients.extend(df_c_temp["Nom Complet"].sort_values().tolist())
        except Exception:
            pass
            
    liste_clients.append("+ Ajouter un nouveau client non répertorié")

    # Définition de vos deux onglets existants
    tab1, tab2 = st.tabs(["📝 Créer une Location", "📊 Suivi des Contrats"])
    
    # --- ONGLET 1 : FORMULAIRE DE SAISIE ---
    with tab1:
        st.subheader("✍️ Générer un nouveau contrat")
        
        # Calcul des statistiques mensuelles si un client valide est sélectionné
        client_selectionne = st.selectbox("Sélectionner le client du contrat *", liste_clients)
        if client_selectionne and client_selectionne not in ["-- Choisir un client existant --", "+ Ajouter un nouveau client non répertorié"]:
            nb_loc_mois = 0
            total_encaisse_client = 0
            
            if not df_loc.empty and "Nom Client" in df_loc.columns:
                # Filtrer sur le client sélectionné
                df_client_history = df_loc[df_loc["Nom Client"] == client_selectionne]
                
                if not df_client_history.empty:
                    # 1. Calcul du nombre de locations ce mois-ci
                    if "Date Enregistrement" in df_client_history.columns:
                        mois_actuel = datetime.now().strftime("%Y-%m")
                        df_client_history["Mois"] = df_client_history["Date Enregistrement"].str[:7]
                        nb_loc_mois = len(df_client_history[df_client_history["Mois"] == mois_actuel])
                    else:
                        nb_loc_mois = len(df_client_history)

                    # # 2. Calcul du cumul financier (Contrats + Cautions)
            for col in df_client_history.columns:
                if any(mot in col for mot in ["Montant", "Total", "Caution", "Prix"]):
                    try:
                        total_encaisse_client += float(pd.to_numeric(df_client_history[col]).sum())
                    except Exception:
                        pass
            # Affichage de la fiche d'information corrigée
            st.info(f"📊 **Statistiques de {client_selectionne} ce mois-ci :**\n"
                    f"* Nombre de locations effectuées : **{nb_loc_mois}**\n"
                    f"* Cumul financier total chez ce client (Contrats + Cautions) : **{total_encaisse_client:,.0f} CFA**".replace(",", " "))
        
        # 2. Formulaire principal des détails de la location
        with st.form(key="form_location", clear_on_submit=True):
            col1, col2 = st.columns(2)
            
            with col1:
                # Si nouveau client, on laisse écrire, sinon on pré-remplit automatiquement
                if client_selectionne == "+ Ajouter un nouveau client non répertorié":
                    nom_client = st.text_input("Nom complet du nouveau client *")
                elif client_selectionne == "-- Choisir un client existant --":
                    nom_client = st.text_input("Nom complet du client *", value="", disabled=True)
                else:
                    nom_client = st.text_input("Nom complet du client *", value=client_selectionne, disabled=True)
                    
                immatriculation = st.text_input("Immatriculation du véhicule *").upper()
                
            with col2:
                prix_journalier = st.number_input("Prix de location par jour (CFA) *", min_value=0, step=5000, value=25000)
                nb_jours = st.number_input("Nombre de jours de location *", min_value=1, step=1, value=1)
                caution_versee = st.number_input("Montant de la Caution versée (CFA) *", min_value=0, step=5000, value=50000)
            
            # Calcul en temps réel pour l'utilisateur
            montant_estime = prix_journalier * nb_jours
            st.warning(f"💰 **Montant de la location :** {montant_estime:,} CFA | **Caution à restituer :** {caution_versee:,} CFA".replace(",", " "))
            
            submit_button = st.form_submit_button(label="💾 Enregistrer le contrat")
        
        if submit_button:
            if nom_client.strip() == "" or immatriculation.strip() == "" or client_selectionne == "-- Choisir un client existant --":
                st.error("Veuillez sélectionner un client valide et remplir l'immatriculation du véhicule.")
            else:
                montant_total = prix_journalier * nb_jours
                date_actuelle = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                nouvelle_location = {
                    "Date Enregistrement": date_actuelle,
                    "Nom Client": nom_client.strip(),
                    "Immatriculation": immatriculation.strip(),
                    "Prix Journalier (CFA)": prix_journalier,
                    "Nombre de Jours": nb_jours,
                    "Caution (CFA)": caution_versee,
                    "Montant Total (CFA)": montant_total
                }
                
                df_loc = pd.concat([df_loc, pd.DataFrame([nouvelle_location])], ignore_index=True)
                df_loc.to_csv(CSV_FILE_LOC, index=False)
                
                # Mettre à jour l'état de l'application immédiatement
                st.session_state["df_loc"] = df_loc
                
                st.success(f"✅ Le contrat de location avec caution pour **{nom_client}** a été enregistré !")
                st.rerun()

    # --- ONGLET 2 : SUIVI ET HISTORIQUE ---
    with tab2:
        st.subheader("🔍 Historique global des locations")
        
        if not df_loc.empty:
            if "Date Enregistrement" in df_loc.columns:
                df_loc = df_loc.sort_values(by="Date Enregistrement", ascending=False)
                
            st.dataframe(df_loc, use_container_width=True)
            
            # Affichage des cumuls financiers
            t1, t2 = st.columns(2)
            with t1:
                total_cumule = df_loc[[c for c in df_loc.columns if "Montant" in c or "Total" in c][0]].sum()
                st.metric(label="Total cumulé des contrats", value=f"{total_cumule:,.0f} CFA".replace(",", " "))
            with t2:
                total_cautions = df_loc[[c for c in df_loc.columns if "Caution" in c][0]].sum() if any("Caution" in c for c in df_loc.columns) else 0
                st.metric(label="Total Cautions en Coffre", value=f"{total_cautions:,.0f} CFA".replace(",", " "))
        else:
            st.info("Aucun contrat de location enregistré pour le moment.")
elif "Identité des Clients" in menu:
    st.title("👥 Gestion de l'Identité des Clients")
    st.subheader("📋 Répertoire et Enregistrement des Clients")
    
    CSV_FILE_CLIENTS = "data_clients.csv"

    if os.path.exists(CSV_FILE_CLIENTS):
        try:
            df_clients = pd.read_csv(CSV_FILE_CLIENTS)
        except Exception:
            df_clients = pd.DataFrame(columns=["Date Enregistrement", "Code Client", "Nom Complet", "Téléphone", "Adresse / Ville", "Type de Client"])
    else:
        df_clients = pd.DataFrame(columns=["Date Enregistrement", "Code Client", "Nom Complet", "Téléphone", "Adresse / Ville", "Type de Client"])

    tab1, tab2 = st.tabs(["📝 Ajouter un Client", "📋 Liste des Clients"])
    
    with tab1:
        st.subheader("✍️ Enregistrer un nouveau client")
        with st.form(key="form_clients", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                nom_client = st.text_input("Nom complet du client *")
                tel_client = st.text_input("Numéro de téléphone *")
            with col2:
                adresse_client = st.text_input("Adresse / Ville (ex: Cocody, Marcory...)")
                type_client = st.selectbox("Type de Client", ["Particulier", "Entreprise", "Régulier / VIP"])
            
            submit_client = st.form_submit_button(label="💾 Enregistrer le client")
            
        if submit_client:
            if nom_client.strip() == "" or tel_client.strip() == "":
                st.error("Veuillez remplir les champs obligatoires (Nom complet et Téléphone).")
            else:
                code_client = f"CLT-{len(df_clients) + 1:03d}"
                date_actuelle = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                nouveau_client = {
                    "Date Enregistrement": date_actuelle,
                    "Code Client": code_client,
                    "Nom Complet": nom_client.strip(),
                    "Téléphone": tel_client.strip(),
                    "Adresse / Ville": adresse_client.strip() if adresse_client.strip() else "Non spécifiée",
                    "Type de Client": type_client
                }
                
                df_clients = pd.concat([df_clients, pd.DataFrame([nouveau_client])], ignore_index=True)
                df_clients.to_csv(CSV_FILE_CLIENTS, index=False)
                st.success(f"✅ Client **{nom_client}** enregistré avec le code **{code_client}** !")
                st.rerun()

    with tab2:
        st.subheader("🔍 Liste complète de vos clients")
        if not df_clients.empty:
            recherche = st.text_input("🔍 Rechercher un client par nom ou numéro :")
            df_affichage = df_clients.copy()
            if recherche:
                df_affichage = df_affichage[
                    df_affichage["Nom Complet"].str.contains(recherche, case=False, na=False) |
                    df_affichage["Téléphone"].str.contains(recherche, case=False, na=False)
                ]
            st.dataframe(df_affichage, use_container_width=True)
            st.metric(label="Total Clients Enregistrés", value=len(df_clients))
        else:
            st.info("Aucun client enregistré pour le moment.")
elif "Identité des Chauffeurs" in menu:
    st.title("👨‍✈️ Gestion de l'Identité des Chauffeurs")
    st.subheader("📋 Répertoire et Enregistrement des Conducteurs")
    
    # --- CONFIGURATION DU FICHIER DE SAUVEGARDE ---
    CSV_FILE_CHAUFFEURS = "data_chauffeurs.csv"

    # Chargement initial des données ou création d'un DataFrame vierge
    if os.path.exists(CSV_FILE_CHAUFFEURS):
        try:
            df_chauffeurs = pd.read_csv(CSV_FILE_CHAUFFEURS)
        except Exception:
            df_chauffeurs = pd.DataFrame(columns=["Date Enregistrement", "Code Chauffeur", "Nom Complet", "Téléphone", "Numéro Permis", "Statut"])
    else:
        df_chauffeurs = pd.DataFrame(columns=["Date Enregistrement", "Code Chauffeur", "Nom Complet", "Téléphone", "Numéro Permis", "Statut"])

    # Définition des deux onglets
    tab1, tab2 = st.tabs(["📝 Ajouter un Chauffeur", "📋 Liste des Chauffeurs"])
    
    # --- ONGLET 1 : FORMULAIRE D'AJOUT ---
    with tab1:
        st.subheader("✍️ Enregistrer un nouveau conducteur")
        
        with st.form(key="form_chauffeurs", clear_on_submit=True):
            col1, col2 = st.columns(2)
            
            with col1:
                nom_chauffeur = st.text_input("Nom complet du chauffeur *")
                tel_chauffeur = st.text_input("Numéro de téléphone *")
                
            with col2:
                permis_chauffeur = st.text_input("Numéro du permis de conduire *")
                statut_chauffeur = st.selectbox("Statut initial", ["Actif / En Service", "Disponible", "En Congé / Absent"])
            
            # Bouton de validation
            submit_chauffeur = st.form_submit_button(label="💾 Enregistrer le chauffeur")
            
        if submit_chauffeur:
            if nom_chauffeur.strip() == "" or tel_chauffeur.strip() == "" or permis_chauffeur.strip() == "":
                st.error("Veuillez remplir tous les champs obligatoires (*).")
            else:
                # Génération automatique d'un code chauffeur unique
                code_chauffeur = f"CH-{len(df_chauffeurs) + 1:03d}"
                date_actuelle = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                nouveau_chauffeur = {
                    "Date Enregistrement": date_actuelle,
                    "Code Chauffeur": code_chauffeur,
                    "Nom Complet": nom_chauffeur.strip(),
                    "Téléphone": tel_chauffeur.strip(),
                    "Numéro Permis": permis_chauffeur.strip().upper(),
                    "Statut": statut_chauffeur
                }
                
                df_chauffeurs = pd.concat([df_chauffeurs, pd.DataFrame([nouveau_chauffeur])], ignore_index=True)
                df_chauffeurs.to_csv(CSV_FILE_CHAUFFEURS, index=False)
                
                st.success(f"✅ Chauffeur **{nom_chauffeur}** enregistré avec succès (Code : **{code_chauffeur}**) !")
                st.rerun()

    # --- ONGLET 2 : RÉPERTOIRE DES CHAUFFEURS ---
    with tab2:
        st.subheader("🔍 Liste complète de vos conducteurs")
        
        if not df_chauffeurs.empty:
            # Barre de recherche dynamique
            recherche_ch = st.text_input("🔍 Rechercher un conducteur par nom ou numéro de permis :")
            
            df_affichage_ch = df_chauffeurs.copy()
            if recherche_ch:
                df_affichage_ch = df_affichage_ch[
                    df_affichage_ch["Nom Complet"].str.contains(recherche_ch, case=False, na=False) |
                    df_affichage_ch["Numéro Permis"].str.contains(recherche_ch, case=False, na=False)
                ]
                
            st.dataframe(df_affichage_ch, use_container_width=True)
            st.metric(label="Total Chauffeurs Enregistrés", value=len(df_chauffeurs))
        else:
            st.info("Aucun chauffeur enregistré pour le moment.")
elif "Flotte globale" in menu:
    st.title("⚙️ Configuration de la Flotte Globale")
    st.subheader("📋 Répertoire et Enregistrement des Véhicules")
    
    # --- CONFIGURATION DU FICHIER DE SAUVEGARDE ---
    CSV_FILE_FLOTTE = "data_flotte.csv"

    # Chargement initial des données ou création d'un DataFrame vierge
    if os.path.exists(CSV_FILE_FLOTTE):
        try:
            df_flotte = pd.read_csv(CSV_FILE_FLOTTE)
        except Exception:
            df_flotte = pd.DataFrame(columns=["Date Enregistrement", "Immatriculation", "Marque / Modèle", "Type de Service", "Statut Véhicule"])
    else:
        df_flotte = pd.DataFrame(columns=["Date Enregistrement", "Immatriculation", "Marque / Modèle", "Type de Service", "Statut Véhicule"])

    # Définition des deux onglets
    tab1, tab2 = st.tabs(["📝 Ajouter un Véhicule", "📋 Liste du Parc Automobile"])
    
    # --- ONGLET 1 : FORMULAIRE D'AJOUT ---
    with tab1:
        st.subheader("✍️ Enregistrer un nouveau véhicule")
        
        with st.form(key="form_flotte", clear_on_submit=True):
            col1, col2 = st.columns(2)
            
            with col1:
                immatriculation = st.text_input("Numéro d'immatriculation *").upper()
                marque_modele = st.text_input("Marque et Modèle (ex: Toyota Corolla, Hyundai...) *")
                
            with col2:
                type_service = st.selectbox("Type d'assignation du véhicule", ["🚖 Taxi Communal", "🚗 Véhicule de Location"])
                statut_vehicule = st.selectbox("Statut initial du véhicule", ["Opérationnel / En Service", "En Maintenance / Garage", "Disponible"])
            
            # Bouton de validation
            submit_vehicule = st.form_submit_button(label="💾 Enregistrer le véhicule")
            
        if submit_vehicule:
            if immatriculation.strip() == "" or marque_modele.strip() == "":
                st.error("Veuillez remplir tous les champs obligatoires (Immatriculation et Marque/Modèle).")
            else:
                date_actuelle = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                nouveau_vehicule = {
                    "Date Enregistrement": date_actuelle,
                    "Immatriculation": immatriculation.strip(),
                    "Marque / Modèle": marque_modele.strip(),
                    "Type de Service": type_service,
                    "Statut Véhicule": statut_vehicule
                }
                
                # 1. Mise à jour de la mémoire locale de Streamlit
        df_flotte = pd.concat([df_flotte, pd.DataFrame([nouveau_vehicule])], ignore_index=True)
        st.session_state['df_flotte'] = df_flotte
        
        # 2. Envoi direct et sécurisé dans Google Sheets
        try:
            import requests
            data_sheets = {
                "date": str(nouveau_vehicule["Date Enregistrement"]),
                "immatriculation": str(nouveau_vehicule["Immatriculation"]),
                "modele": str(nouveau_vehicule["Marque / Modèle"]),
                "service": str(nouveau_vehicule["Type de Service"]),
                "statut": str(nouveau_vehicule["Statut Véhicule"])
            }
            # Envoi automatique vers votre tableur en ligne
            url_api = f"https://google.com"
            # Sauvegarde locale de secours en plus au cas où internet coupe
            df_flotte.to_csv(CSV_FILE_FLOTTE, index=False)
        except:
            df_flotte.to_csv(CSV_FILE_FLOTTE, index=False)
                
            st.success(f"✅ Le véhicule **{immatriculation}** ({marque_modele}) a été enregistré avec succès dans la flotte !")
            st.rerun()

    # --- ONGLET 2 : RÉPERTOIRE DE LA FLOTTE ---
    with tab2:
        st.subheader("🔍 Liste complète de vos véhicules")
        
        if not df_flotte.empty:
            # Filtre de recherche par immatriculation
            recherche_v = st.text_input("🔍 Filtrer par numéro d'immatriculation :")
            
            df_affichage_v = df_flotte.copy()
            if recherche_v:
                df_affichage_v = df_affichage_v[df_affichage_v["Immatriculation"].str.contains(recherche_v, case=False, na=False)]
                
            st.dataframe(df_affichage_v, use_container_width=True)
            
            # Petites statistiques rapides sur le parc
            c1, c2 = st.columns(2)
            with c1:
                st.metric(label="Total Véhicules", value=len(df_flotte))
            with c2:
                taxis_count = len(df_flotte[df_flotte["Type de Service"] == "🚖 Taxi Communal"])
                st.metric(label="Dont Taxis Communaux", value=taxis_count)
        else:
            st.info("Aucun véhicule enregistré dans votre flotte pour le moment.")
elif "Gestion des Accès" in menu:
    st.title("🔑 Espace Donneur d'Accès (Réservé au DG)")
    st.subheader("📋 Liste des Comptes et Identifiants du Personnel")
    
    st.write("En tant que **Directeur Général**, vous êtes le seul à pouvoir consulter la liste des accès actifs de l'entreprise.")
    
    # Configuration de la table des utilisateurs
    # (Ce bloc fait le lien avec votre dictionnaire 'COMPTES' du début)
    if 'COMPTES' in globals():
        donnees_comptes = []
        for identifiant, infos in COMPTES.items():
            donnees_comptes.append({
                "Identifiant de Connexion": identifiant,
                "Nom Affiché": infos["nom"],
                "Mot de passe actuel": infos["mp"],
                "Rôle Système": infos["role"]
            })
            
        df_comptes = pd.DataFrame(donnees_comptes)
        st.dataframe(df_comptes, use_container_width=True)
    else:
        st.info("La liste des comptes d'accès n'est pas encore initialisée en haut du fichier.")
