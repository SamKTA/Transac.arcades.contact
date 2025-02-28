import streamlit as st
import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ID de la feuille Google Sheets pour les contacts
CONTACT_SHEET_ID = "1f0x47zQrmCdo9GwF_q2wTOiP9jxEvMmLevY7xmDOp4A"

# Configuration de la page
st.set_page_config(
    page_title="Gestion Contacts ORPI Arcades",
    page_icon=":telephone:",
    initial_sidebar_state="collapsed",
    layout="wide"
)

# Initialiser les variables de session
if "page" not in st.session_state:
    st.session_state.page = "roulement"

if "conseiller_selectionne" not in st.session_state:
    st.session_state.conseiller_selectionne = None

if "type_roulement" not in st.session_state:
    st.session_state.type_roulement = None

if "formulaire_soumis" not in st.session_state:
    st.session_state.formulaire_soumis = False

# Liste des conseillers et leurs noms complets
CONSEILLERS = {
    "Catherine": "Catherine DUPONT",  # À remplacer par les noms réels
    "Antoine": "Antoine MARTIN",
    "Mélissa": "Mélissa ROUX",
    "Nicolas": "Nicolas BERNARD",
    "Naomi": "Naomi PETIT",
    "Fayçal": "Fayçal DURAND",
    "Damien": "Damien LEROY",
    "Mathilde": "Mathilde MOREAU"
}

# Liste des emails des conseillers
EMAILS_CONSEILLERS = {
    "Catherine DUPONT": "catherine@orpi.com",
    "Antoine MARTIN": "antoine@orpi.com",
    "Mélissa ROUX": "melissa@orpi.com",
    "Nicolas BERNARD": "nicolas@orpi.com",
    "Naomi PETIT": "naomi@orpi.com",
    "Fayçal DURAND": "faycal@orpi.com",
    "Damien LEROY": "damien@orpi.com",
    "Mathilde MOREAU": "mathilde@orpi.com",
    "Clément VIGREUX": "clement.vigreux@orpi.com",
    "Pascal BOFFERON": "pascal.bofferon@orpi.com",
    "Angélique CHENERAILLES": "angélique.chenerailles@orpi.com",
    "Bertrand FOURNIER": "bertrand.fournier.agencedesarcades@orpi.com",
    "Joshua BESSE": "joshua.besse@orpi.com",
    "Irina GALOYAN": "irina@orpi.com",
    "Arnaud SELLAM": "arnaud.sellam@orpi.com",
    "Benoît COUSTEAUD": "benoît.cousteaud@orpi.com",
    "Orianne BOULESTEIX": "orianne@orpi.com",
    "Cyril REINICHE": "cyrilreiniche@orpi.com",
    "Sam.test": "skita@orpi.com"
}

# Fonction pour se connecter à Google Sheets
def get_sheets_client():
    """Établit une connexion avec l'API Google Sheets"""
    scope = [
        'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive'
    ]
    
    creds_dict = st.secrets["google_credentials"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    
    return gspread.authorize(creds)

# Fonctions pour la gestion du roulement
def lire_roulements():
    """Lit l'état actuel des roulements depuis Google Sheets"""
    try:
        client = get_sheets_client()
        doc = client.open_by_key(ROULEMENT_SHEET_ID)
        
        # Lire les données de la feuille État
        etat_sheet = doc.worksheet("État")
        data = etat_sheet.get_all_records()
        
        # Convertir en DataFrame
        df = pd.DataFrame(data)
        
        # Lire les indisponibilités
        indispo_sheet = doc.worksheet("Indisponibilités")
        indispo_data = indispo_sheet.get_all_records()
        indispo_df = pd.DataFrame(indispo_data)
        
        return df, indispo_df
    except Exception as e:
        st.error(f"Erreur lors de la lecture des roulements : {e}")
        # Créer des données par défaut si erreur
        default_df = pd.DataFrame({
            "Type": ["VENDEURS PROJET VENTE", "ACQUÉREURS", "VENDEURS PAS DE PROJET"],
            "Dernier_Conseiller": ["", "", ""]
        })
        default_indispo = pd.DataFrame(columns=["Conseiller", "Début", "Fin", "Raison"])
        return default_df, default_indispo

def mettre_a_jour_roulement(type_roulement, conseiller):
    """Met à jour le roulement dans Google Sheets"""
    try:
        client = get_sheets_client()
        doc = client.open_by_key(ROULEMENT_SHEET_ID)
        
        # Mettre à jour l'état
        etat_sheet = doc.worksheet("État")
        
        # Trouver la ligne correspondant au type de roulement
        cellule = etat_sheet.find(type_roulement)
        row = cellule.row
        
        # Mettre à jour le dernier conseiller
        etat_sheet.update_cell(row, 2, conseiller)
        
        # Ajouter à l'historique
        historique_sheet = doc.worksheet("Historique")
        now = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        historique_sheet.append_row([now, type_roulement, conseiller])
        
        return True
    except Exception as e:
        st.error(f"Erreur lors de la mise à jour du roulement : {e}")
        return False

def est_disponible(conseiller, indispo_df):
    """Vérifie si un conseiller est disponible aujourd'hui"""
    if indispo_df.empty:
        return True
    
    aujourd_hui = datetime.now().date()
    
    for _, row in indispo_df.iterrows():
        if row["Conseiller"] == conseiller:
            try:
                debut = datetime.strptime(row["Début"], "%d/%m/%Y").date()
                fin = datetime.strptime(row["Fin"], "%d/%m/%Y").date()
                
                if debut <= aujourd_hui <= fin:
                    return False
            except ValueError:
                # Si la date n'est pas au bon format, on continue
                continue
    
    return True

def obtenir_prochain_conseiller(type_roulement, dernier_conseiller, indispo_df):
    """Détermine le prochain conseiller dans le roulement"""
    if type_roulement == "VENDEURS PROJET VENTE":
        order = ["Catherine", "Antoine", "Mélissa", "Nicolas", "Naomi", "Fayçal", "Damien", "Mathilde"]
    elif type_roulement == "ACQUÉREURS":
        order = ["Catherine", "Antoine", "Mélissa", "Nicolas", "Naomi", "Fayçal", "Damien", "Mathilde"]
    elif type_roulement == "VENDEURS PAS DE PROJET":
        order = ["Catherine", "Damien", "Mathilde", "Naomi", "Nicolas", "Fayçal", "Antoine", "Mélissa"]
    else:
        return None
    
    # Si pas de dernier conseiller, commencer par le premier
    if not dernier_conseiller or dernier_conseiller not in order:
        idx = 0
    else:
        # Trouver l'index du dernier conseiller
        idx = order.index(dernier_conseiller)
        # Passer au suivant
        idx = (idx + 1) % len(order)
    
    # Vérifier la disponibilité et parcourir la liste
    start_idx = idx
    while True:
        conseiller = order[idx]
        
        if est_disponible(conseiller, indispo_df):
            return conseiller
        
        # Passer au conseiller suivant
        idx = (idx + 1) % len(order)
        
        # Si on a fait le tour complet, retourner le premier même si indisponible
        if idx == start_idx:
            return order[idx]

# Fonctions pour les emails et Google Sheets
def sauvegarder_dans_sheets(donnees):
    """
    Sauvegarde les données du formulaire dans un Google Sheet
    dans la feuille "All" et dans la feuille correspondant au conseiller
    """
    try:
        # Configuration pour l'accès à Google Sheets
        client = get_sheets_client()
        
        # Ouvrir le document
        sheet_doc = client.open_by_key(CONTACT_SHEET_ID)
        
        # Préparer les données à enregistrer (dans l'ordre des colonnes du Google Sheet)
        row_data = [
            donnees["date"],                # A - Date réception du contact
            donnees["assistante"],          # B - Assistante
            donnees["destinataire"],        # C - Conseiller
            donnees["source"],              # D - Source
            donnees["canal"],               # E - Canal
            donnees["type_contact"],        # F - Type contact
            donnees["nom_client"],          # G - Nom complet du client
            donnees["email_client"],        # H - Adresse e-mail
            donnees["telephone_client"],    # I - Téléphone
            donnees["commentaire"]          # J - Commentaire
        ]
        
        # 1. Sauvegarde dans la feuille "All"
        try:
            all_sheet = sheet_doc.worksheet("All")
            all_sheet.append_row(row_data)
        except Exception as e:
            st.warning(f"Impossible d'ajouter à la feuille 'All': {e}")
        
        # 2. Sauvegarde dans la feuille du conseiller
        try:
            # Utiliser la convention de nommage exacte des feuilles (prénom du conseiller)
            conseiller_sheet_name = donnees["destinataire"].split()[0]  # Prend juste le prénom
            # Tentative d'accès à la feuille du conseiller
            conseiller_sheet = sheet_doc.worksheet(conseiller_sheet_name)
            conseiller_sheet.append_row(row_data)
        except gspread.exceptions.WorksheetNotFound:
            st.warning(f"Feuille '{conseiller_sheet_name}' non trouvée. Les données sont uniquement sauvegardées dans 'All'.")
        except Exception as e:
            st.warning(f"Impossible d'ajouter à la feuille '{conseiller_sheet_name}': {e}")
            
        return True
    except Exception as e:
        st.error(f"Erreur lors de la sauvegarde dans Google Sheets : {e}")
        return False

def send_email(receiver_email, email_data):
    """
    Fonction pour envoyer un email
    """
    try:
        # Configuration du serveur SMTP
        smtp_server = "smtp.gmail.com"
        port = 587
        # Adresse email de l'expéditeur
        sender_email = "contactpro.skdigital@gmail.com"
        # Mot de passe d'application Gmail
        password = "nxti raqi vwwu zvng"

        # Préparation du contenu de l'email selon le template
        email_content = f"""Bonjour {email_data['destinataire']}, 

Un contact {email_data['type_contact'].lower()} souhaite que tu le recontactes ou que tu confirmes un rendez-vous avec lui (à vérifier dans les commentaires). Voici le commentaire : 

{email_data['commentaire']}

Voici ces coordonnées : CONTACT {email_data['type_contact'].upper()}

{email_data['nom_client']}
{email_data['email_client']}
{email_data['telephone_client']}

Bon appel & bonne journée à toi !
"""

        # Créer le message
        message = MIMEMultipart()
        message["From"] = f"Transmission Contact ORPI Arcades <{sender_email}>"
        message["To"] = receiver_email
        message["Subject"] = f"🔔 +1 CONTACT : {email_data['nom_client']} - {email_data['type_contact'].upper()}"
        message.attach(MIMEText(email_content, "plain"))

        # Établir la connexion SMTP et envoyer l'email
        with smtplib.SMTP(smtp_server, port) as server:
            server.starttls()
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_email, message.as_string())
        
        return True
    except smtplib.SMTPAuthenticationError:
        st.error("Erreur d'authentification. Vérifiez vos identifiants.")
        return False
    except Exception as e:
        st.error(f"Erreur lors de l'envoi de l'email : {e}")
        return False

def get_destinataire_email(destinataire):
    """
    Retourne l'adresse email en fonction du destinataire sélectionné
    """
    return EMAILS_CONSEILLERS.get(destinataire, "")

# Page de roulement
def page_roulement():
    """Affiche la page de gestion des roulements"""
    st.title("Roulement des contacts ORPI Arcades")
    
    # Lire l'état des roulements
    roulements_df, indispo_df = lire_roulements()
    
    # Afficher les trois options de roulement
    st.header("Sélectionner le type de contact")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.subheader("VENDEURS PROJET VENTE")
        vendeur_projet = roulements_df[roulements_df["Type"] == "VENDEURS PROJET VENTE"]
        dernier_vendeur_projet = vendeur_projet["Dernier_Conseiller"].values[0] if not vendeur_projet.empty else ""
        
        prochain = obtenir_prochain_conseiller("VENDEURS PROJET VENTE", dernier_vendeur_projet, indispo_df)
        
        st.write(f"Dernier contact attribué à: **{dernier_vendeur_projet}**")
        st.write(f"Prochain contact à attribuer à: **{prochain}**")
        
        if st.button("Sélectionner VENDEURS PROJET VENTE", key="btn_vendeur_projet"):
            st.session_state.type_roulement = "VENDEURS PROJET VENTE"
            st.session_state.conseiller_selectionne = CONSEILLERS.get(prochain, prochain)
            st.session_state.page = "formulaire"
            st.experimental_rerun()
    
    with col2:
        st.subheader("ACQUÉREURS")
        acquereur = roulements_df[roulements_df["Type"] == "ACQUÉREURS"]
        dernier_acquereur = acquereur["Dernier_Conseiller"].values[0] if not acquereur.empty else ""
        
        prochain = obtenir_prochain_conseiller("ACQUÉREURS", dernier_acquereur, indispo_df)
        
        st.write(f"Dernier contact attribué à: **{dernier_acquereur}**")
        st.write(f"Prochain contact à attribuer à: **{prochain}**")
        
        if st.button("Sélectionner ACQUÉREURS", key="btn_acquereur"):
            st.session_state.type_roulement = "ACQUÉREURS"
            st.session_state.conseiller_selectionne = CONSEILLERS.get(prochain, prochain)
            st.session_state.page = "formulaire"
            st.experimental_rerun()
    
    with col3:
        st.subheader("VENDEURS PAS DE PROJET")
        vendeur_pas_projet = roulements_df[roulements_df["Type"] == "VENDEURS PAS DE PROJET"]
        dernier_vendeur_pas_projet = vendeur_pas_projet["Dernier_Conseiller"].values[0] if not vendeur_pas_projet.empty else ""
        
        prochain = obtenir_prochain_conseiller("VENDEURS PAS DE PROJET", dernier_vendeur_pas_projet, indispo_df)
        
        st.write(f"Dernier contact attribué à: **{dernier_vendeur_pas_projet}**")
        st.write(f"Prochain contact à attribuer à: **{prochain}**")
        
        if st.button("Sélectionner VENDEURS PAS DE PROJET", key="btn_vendeur_pas_projet"):
            st.session_state.type_roulement = "VENDEURS PAS DE PROJET"
            st.session_state.conseiller_selectionne = CONSEILLERS.get(prochain, prochain)
            st.session_state.page = "formulaire"
            st.experimental_rerun()
    
    # Section pour la gestion des indisponibilités
    st.header("Gestion des indisponibilités")
    
    with st.expander("Voir/Modifier les indisponibilités"):
        # Afficher les indisponibilités actuelles
        if not indispo_df.empty:
            st.subheader("Indisponibilités actuelles")
            st.dataframe(indispo_df)
        
        # Formulaire pour ajouter une indisponibilité
        st.subheader("Ajouter une indisponibilité")
        with st.form("form_indispo"):
            conseiller = st.selectbox("Conseiller", list(CONSEILLERS.keys()))
            date_debut = st.date_input("Date de début")
            date_fin = st.date_input("Date de fin")
            raison = st.text_input("Raison (optionnelle)")
            
            submitted = st.form_submit_button("Ajouter")
            
            if submitted:
                try:
                    client = get_sheets_client()
                    doc = client.open_by_key(ROULEMENT_SHEET_ID)
                    indispo_sheet = doc.worksheet("Indisponibilités")
                    
                    indispo_sheet.append_row([
                        conseiller,
                        date_debut.strftime("%d/%m/%Y"),
                        date_fin.strftime("%d/%m/%Y"),
                        raison
                    ])
                    
                    st.success(f"Indisponibilité ajoutée pour {conseiller}")
                    st.experimental_rerun()
                except Exception as e:
                    st.error(f"Erreur lors de l'ajout de l'indisponibilité : {e}")

# Page du formulaire de contact
def page_formulaire():
    """Affiche la page du formulaire de contact"""
    st.title("Transmission contact ORPI Arcades")
    st.subheader("Formulaire de transmission")
    
    # Date du jour automatique
    date_aujourd_hui = datetime.now().strftime("%d/%m/%Y")
    st.write(f"**Date :** {date_aujourd_hui}")
    
    # Bouton retour
    if st.button("← Retour au roulement"):
        st.session_state.page = "roulement"
        st.experimental_rerun()
    
    # Formulaire de saisie
    with st.form(key='formulaire_contact'):
        # Assistante
        assistante = st.selectbox("Assistante", 
                             options=["Laura", "Léonor"])
        
        # Destinataire (avec pré-sélection si venant du roulement)
        destinataire_options = list(EMAILS_CONSEILLERS.keys())
        
        if st.session_state.conseiller_selectionne:
            try:
                preselection_index = destinataire_options.index(st.session_state.conseiller_selectionne)
            except ValueError:
                preselection_index = 0
                
            destinataire = st.selectbox("Ce contact est pour", 
                                       options=destinataire_options,
                                       index=preselection_index)
        else:
            destinataire = st.selectbox("Ce contact est pour", 
                                       options=destinataire_options)
        
        # Source
        source = st.selectbox("Source", 
                             options=["LBC", "SeLoger", "SAO", "Prospection", 
                                      "Notoriété", "Recommandation", "Réseaux sociaux"])
        
        # Canal
        canal = st.selectbox("Canal", 
                            options=["Appel téléphonique", "Passage agence"])
        
        # Type contact
        type_contact = st.selectbox("Type contact", 
                                   options=["Acheteur", "Vendeur", "Acheteur mail SB"])
        
        # Nom complet du client (obligatoire)
        nom_client = st.text_input("Nom complet du client *", placeholder="Nom et prénom")
        
        # Adresse e-mail (optionnel)
        email_client = st.text_input("Adresse e-mail", placeholder="Email du client")
        
        # Téléphone client (obligatoire)
        telephone_client = st.text_input("Téléphone *", placeholder="Numéro de téléphone")
        
        # Commentaire
        commentaire = st.text_area("Commentaire", placeholder="Détails supplémentaires")
        
        # Bouton de validation
        submitted = st.form_submit_button("Je valide")
        
        # Validation du formulaire
        if submitted:
            # Vérification des champs obligatoires
            if not telephone_client or not nom_client:
                st.error("Merci de remplir tous les champs obligatoires (*)")
            else:
                # Déterminer l'adresse email du destinataire
                email_destinataire = get_destinataire_email(destinataire)
                
                # Préparer un dictionnaire avec les données
                donnees = {
                    "date": date_aujourd_hui,
                    "assistante": assistante,
                    "destinataire": destinataire,
                    "source": source,
                    "canal": canal,
                    "type_contact": type_contact,
                    "nom_client": nom_client,
                    "email_client": email_client,
                    "telephone_client": telephone_client,
                    "commentaire": commentaire
                }
                
                # Sauvegarder dans Google Sheets
                if sauvegarder_dans_sheets(donnees):
                    # Envoi de l'email
                    if send_email(email_destinataire, donnees):
                        st.success(f"Contact transmis avec succès à {destinataire} !")
                        
                        # Mettre à jour le roulement si on vient de la page roulement
                        if st.session_state.type_roulement:
                            # Récupérer le prénom du conseiller
                            prenom = destinataire.split()[0]
                            
                            # Chercher la clé correspondant au nom complet
                            for key, value in CONSEILLERS.items():
                                if value == destinataire:
                                    prenom = key
                                    break
                            
                            mettre_a_jour_roulement(st.session_state.type_roulement, prenom)
                            
                            # Retourner à la page de roulement
                            st.session_state.formulaire_soumis = True
                            st.session_state.type_roulement = None
                            st.session_state.conseiller_selectionne = None
                            
                            # Message de redirection
                            st.info("Redirection vers la page de roulement...")
                            st.button("Retourner au roulement", on_click=lambda: setattr(st.session_state, "page", "roulement"))
                    else:
                        st.error("Problème lors de l'envoi de l'email")
                else:
                    st.error("Problème lors de la sauvegarde des données")

# Fonction principale
def main():
    """Fonction principale de l'application"""
    # Afficher la page demandée
    if st.session_state.page == "roulement":
        page_roulement()
    elif st.session_state.page == "formulaire":
        page_formulaire()

if __name__ == "__main__":
    main()
