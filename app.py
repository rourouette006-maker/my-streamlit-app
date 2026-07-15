import streamlit as st     #(alias st):créer l'interface web de l'application
import pandas as pd    # manipuler et analyser les données (tableaux, CSV, etc.)
import os   #importer le module os pour interagir avec le système d'exploitation (chemins, fichiers, dossiers)
import base64  # Bibliothèque standard pour encoder l'image du logo

# Configuration de la page (et mise en page centé)
st.set_page_config(page_title="Connexion - SMQ", layout="centered")

# --- FONCTION POUR METTRE LE LOGO EN ARRIÈRE-PLAN ---
#(encoder une image locale en Base64 et l'appliquer en arrière-plan avec du CSS personnalisé)
def add_bg_from_local(image_file):
    if os.path.exists(image_file):
        with open(image_file, "rb") as image_file:
            encoded_string = base64.b64encode(image_file.read()).decode()
        st.markdown("<h1 style='text-align: center;'>🔒 Accès Sécurisé - SMQ</h1>", unsafe_allow_html=True)
with st.form("login_form"):
    identifiant = st.text_input("Identifiant")
    password = st.text_input("Mot de passe", type="password")
    submit_button = st.form_submit_button("Se connecter")
        
# --- GESTION DE LA SESSION DE CONNEXION ---
if "connected" not in st.session_state:
    st.session_state["connected"] = False    # Initialise l'état de connexion de la session à "False" s'il n'existe pas encore


# ==========================================
# PAGE 1 : ÉCRAN DE CONNEXION
# ==========================================
if not st.session_state["connected"]:
    # On applique le logo en arrière-plan
    add_bg_from_local("logo.png")
    
    st.title("🔒 Accès Sécurisé - SMQ")
    st.write("Veuillez saisir vos identifiants pour accéder au système de management de la qualité.")    #Affiche l'écran de connexion sécurisé (avec logo en arrière-plan) si l'utilisateur n'est pas connecté

    # Formulaire de connexion (il s'affichera par-dessus le fond transparent)
    with st.form("login_form"):
        username = st.text_input("Identifiant")
        password = st.text_input("Mot de passe", type="password")
        submit_button = st.form_submit_button("Se connecter")
        
        if submit_button:
            if username == "admin" and password == "qualite2026":
                st.session_state["connected"] = True
                st.success("Connexion réussie ! Chargement...")
                st.rerun()
            else:
                st.error("Identifiant ou mot de passe incorrect ❌")    # Gère le formulaire de connexion, vérifie les identifiants et connecte l'utilisateur si les accès sont corrects




# ==========================================
# PAGE 2 : LE DASHBOARD (S'affiche après connexion)
# ==========================================
else:
    # Le reste de votre code existant pour le dashboard reste identique...
    st.sidebar.title("👤 Espace Connecté")
    st.sidebar.write("Rôle : **Responsable Qualité**")
    
    if st.sidebar.button("🚪 Se déconnecter"):
        st.session_state["connected"] = False
        st.rerun()
        
    st.title("📊 Tableau de Bord - SMQ")
    st.markdown("Bienvenue dans votre espace de suivi de la qualité.")
    st.write("---")
    
    try:
        df = pd.read_excel("database.xlsx", sheet_name="NonConformites")
        st.subheader("🔍 Données actuelles dans Excel")
        st.dataframe(df, use_container_width=True)
    except Exception as e:
        st.error(f"Erreur lors du chargement des données : {e}")      # Gère l'affichage du tableau de bord (sidebar, déconnexion et chargement des données Excel) après une connexion réussie
