import streamlit as st
import smtplib
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime

# ID de la feuille Google Sheets - À REMPLACER par ton ID de Google Sheets
SHEET_ID = "TON_ID_DE_GOOGLE_SHEET"

def sauvegarder_dans_sheets(donnees):
    """
    Sauvegarde les données du formulaire dans un Google Sheet
    """
    try:
        # Configuration pour l'accès à Google Sheets
        scope = [
            'https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive'
        ]
        
        # Utiliser les credentials stockés dans Streamlit Secrets
        creds_dict = st.secrets["google_credentials"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        
        # Authentification et accès au Google Sheet
        client = gspread.authorize(creds)
        
        # Ouvrir la première feuille du document
        sheet = client.open_by_key(SHEET_ID).sheet1
        
        # Préparer les données à enregistrer
        row_data = [
            donnees["date"],                # Date réception du contact
            donnees["assistante"],          # Assistante
            donnees["destinataire"],        # Ce contact est pour
            donnees["source"],              # Source
            donnees["canal"],               # Canal
            donnees["nom_client"],          # Nom complet du client
            donnees["email_client"],        # Adresse e-mail
            donnees["telephone_client"],    # Téléphone
            donnees["commentaire"]          # Commentaire
        ]
        
        # Ajouter une nouvelle ligne dans la feuille
        sheet.append_row(row_data)
        
        return True
    except Exception as e:
        st.error(f"Erreur lors de la sauvegarde dans Google Sheets : {e}")
        return False

def send_email(receiver_email, email_content):
    """
    Fonction pour envoyer un email avec gestion sécurisée des identifiants
    """
    try:
        # Configuration du serveur SMTP
        smtp_server = "smtp.gmail.com"
        port = 587
        # À REMPLACER par ton adresse email et ton mot de passe d'application
        sender_email = "ton.email@gmail.com"
        password = st.secrets["email_password"]

        # Créer le message
        message = MIMEMultipart()
        message["From"] = f"Transmission Contact ORPI <{sender_email}>"
        message["To"] = receiver_email
        message["Subject"] = "Nouveau Contact ORPI"
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
    if destinataire == "Samuel KITA":
        return "skita@orpi.com"
    elif destinataire == "Laurie BLONDEAU":
        return "lblondeau@orpi.com"
    return ""

def main():
    # Configuration de la page
    st.set_page_config(page_title="Transmission Contact ORPI", page_icon=":telephone:")
    
    # Titre principal
    st.title("Transmission contact ORPI")
    st.subheader("Formulaire de transmission")
    
    # Date du jour automatique
    date_aujourd_hui = datetime.now().strftime("%d/%m/%Y")
    st.write(f"**Date :** {date_aujourd_hui}")
    
    # Formulaire de saisie
    with st.form(key='formulaire_contact'):
        # Assistante
        assistante = st.selectbox("Assistante", 
                             options=["Laura", "Léonor"])
        
        # Destinataire
        destinataire = st.selectbox("Ce contact est pour", 
                                   options=["Samuel KITA", "Laurie BLONDEAU"])
        
        # Source
        source = st.selectbox("Source", 
                             options=["LBC", "SeLoger", "SAO", "Prospection", 
                                      "Notoriété", "Recommandation", "Réseaux sociaux"])
        
        # Canal
        canal = st.selectbox("Canal", 
                            options=["Appel téléphonique", "Passage agence"])
        
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
                    "nom_client": nom_client,
                    "email_client": email_client,
                    "telephone_client": telephone_client,
                    "commentaire": commentaire
                }
                
                # Préparer le contenu de l'email
                email_content = f"""Bonjour {destinataire},

Nouveau contact transmis par {assistante} le {date_aujourd_hui}.

Voici les coordonnées du client : 

Nom: {nom_client}
Email: {email_client}
Téléphone: {telephone_client}

Ce contact provient de {source} via {canal}.

Commentaire : {commentaire}

Bonne journée,
"""
                # Sauvegarder dans Google Sheets
                if sauvegarder_dans_sheets(donnees):
                    # Envoi de l'email
                    if send_email(email_destinataire, email_content):
                        st.success(f"Contact transmis avec succès à {destinataire} !")
                    else:
                        st.error("Problème lors de l'envoi de l'email")
                else:
                    st.error("Problème lors de la sauvegarde des données")

if __name__ == "__main__":
    main()
