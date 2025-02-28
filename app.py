import streamlit as st
import smtplib
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime

# ID de la feuille Google Sheets - À REMPLACER par ton ID de Google Sheets
SHEET_ID = "1f0x47zQrmCdo9GwF_q2wTOiP9jxEvMmLevY7xmDOp4A"

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
            donnees["type_contact"],        # Type contact
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

def send_email(receiver_email, email_data):
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
    # Dictionnaire des conseillers avec leurs emails
    conseillers = {
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
        "SAM.test": "skita@orpi.com",
    }
    
    # Retourne l'email correspondant au conseiller sélectionné
    return conseillers.get(destinataire, "")

def main():
    # Configuration de la page
    st.set_page_config(page_title="Transmission Contact ORPI Arcades", page_icon=":telephone:")
    
    # Titre principal
    st.title("Transmission contact ORPI Arcades")
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
                                   options=[
                                       "Clément VIGREUX",
                                       "Pascal BOFFERON",
                                       "Angélique CHENERAILLES",
                                       "Bertrand FOURNIER",
                                       "Joshua BESSE",
                                       "Irina GALOYAN",
                                       "Arnaud SELLAM",
                                       "Benoît COUSTEAUD",
                                       "Orianne BOULESTEIX",
                                       "Cyril REINICHE",
                                       "SAM.test",
                                   ])
        
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
                    else:
                        st.error("Problème lors de l'envoi de l'email")
                else:
                    st.error("Problème lors de la sauvegarde des données")

if __name__ == "__main__":
    main()
