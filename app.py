import streamlit as st
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import time

# ============================================
# CONFIGURATION DE LA PAGE
# ============================================
st.set_page_config(
    page_title="📧 Publipostage Email",
    page_icon="📧",
    layout="wide"
)

st.title("📧 Application de Publipostage")
st.markdown("Importez votre fichier Excel, personnalisez vos emails et envoyez-les en masse !")
st.divider()

# ============================================
# INITIALISATION DE LA SESSION
# ============================================
if 'df' not in st.session_state:
    st.session_state.df = None
if 'signature_bytes' not in st.session_state:
    st.session_state.signature_bytes = None

# ============================================
# FONCTION D'ENVOI D'EMAIL
# ============================================
def envoyer_email(smtp_server, port, sender_email, password,
                   destinataire, sujet, corps_html, signature_bytes=None, use_ssl=False):
    """
    Envoie un email HTML avec signature image optionnelle intégrée.
    """
    try:
        msg = MIMEMultipart('related')
        msg['From'] = sender_email
        msg['To'] = destinataire
        msg['Subject'] = sujet

        # Ajouter la signature image dans le HTML si elle existe
        corps_final = corps_html
        if signature_bytes:
            corps_final += '<br><br><img src="cid:signature_image" style="max-width:300px;">'

        msg_html = MIMEText(corps_final, 'html')
        msg.attach(msg_html)

        # Attacher l'image de signature (format inline)
        if signature_bytes:
            image = MIMEImage(signature_bytes)
            image.add_header('Content-ID', '<signature_image>')
            image.add_header('Content-Disposition', 'inline', filename='signature.png')
            msg.attach(image)

        # Connexion SMTP selon le mode choisi
        if use_ssl:
            server = smtplib.SMTP_SSL(smtp_server, port, timeout=10)
        else:
            server = smtplib.SMTP(smtp_server, port, timeout=10)
            server.starttls()

        server.login(sender_email, password)
        server.sendmail(sender_email, destinataire, msg.as_string())
        server.quit()

        return True, "Email envoyé avec succès ✅"

    except smtplib.SMTPAuthenticationError:
        return False, "❌ Erreur d'authentification : vérifiez votre email et votre mot de passe d'application"
    except smtplib.SMTPConnectError:
        return False, "❌ Impossible de se connecter au serveur SMTP"
    except Exception as e:
        return False, f"❌ Erreur : {str(e)}"


# ============================================
# SECTION 1 : CONFIGURATION SMTP
# ============================================
st.header("1️⃣ Configuration du serveur email")

col1, col2 = st.columns(2)

with col1:
    fournisseur = st.selectbox(
        "📮 Fournisseur email",
        ["Gmail", "Outlook/Hotmail", "Yahoo", "Autre (personnalisé)"]
    )

    # Pré-remplissage selon le fournisseur
    if fournisseur == "Gmail":
        smtp_server_default = "smtp.gmail.com"
        port_default = 587
    elif fournisseur == "Outlook/Hotmail":
        smtp_server_default = "smtp.office365.com"
        port_default = 587
    elif fournisseur == "Yahoo":
        smtp_server_default = "smtp.mail.yahoo.com"
        port_default = 587
    else:
        smtp_server_default = ""
        port_default = 587

    smtp_server = st.text_input("🌐 Serveur SMTP", value=smtp_server_default)
    port = st.number_input("🔌 Port", value=port_default, step=1)

with col2:
    sender_email = st.text_input("📧 Votre adresse email")
    password = st.text_input(
        "🔑 Mot de passe d'application",
        type="password",
        help="Pour Gmail : utilisez un mot de passe d'application (pas votre mot de passe normal)"
    )

use_ssl = (port == 465)
if use_ssl:
    st.info("ℹ️ Port 465 détecté : connexion SSL directe sera utilisée")
else:
    st.info("ℹ️ Connexion STARTTLS sera utilisée")

st.divider()

# ============================================
# SECTION 2 : IMPORT DU FICHIER EXCEL
# ============================================
st.header("2️⃣ Importer votre fichier Excel")

fichier_excel = st.file_uploader(
    "📂 Sélectionnez votre fichier Excel (.xlsx)",
    type=["xlsx", "xls"]
)

if fichier_excel is not None:
    try:
        df = pd.read_excel(fichier_excel)
        st.session_state.df = df
        st.success(f"✅ Fichier importé avec succès : {len(df)} ligne(s) trouvée(s)")

        with st.expander("👀 Aperçu des données"):
            st.dataframe(df, use_container_width=True)

        st.info(f"📋 Colonnes disponibles : {', '.join(df.columns)}")

    except Exception as e:
        st.error(f"❌ Erreur lors de la lecture du fichier : {e}")

st.divider()

# ============================================
# SECTION 3 : SIGNATURE IMAGE
# ============================================
st.header("3️⃣ Signature email (optionnel)")

signature_image = st.file_uploader(
    "🖼️ Importer une image de signature (logo, signature manuscrite, etc.)",
    type=["png", "jpg", "jpeg"],
    help="Cette image sera ajoutée automatiquement à la fin de chaque email"
)

if signature_image is not None:
    st.session_state.signature_bytes = signature_image.getvalue()
    col_a, col_b = st.columns([1, 3])
    with col_a:
        st.image(signature_image, caption="Aperçu", width=200)
    with col_b:
        st.success("✅ Signature chargée ! Elle sera ajoutée à tous les emails envoyés.")
else:
    st.session_state.signature_bytes = None

st.divider()

# ============================================
# SECTION 4 : RÉDACTION DU MESSAGE
# ============================================
st.header("4️⃣ Rédiger votre message")

if st.session_state.df is not None:
    colonnes_disponibles = list(st.session_state.df.columns)
    st.caption(f"💡 Utilisez {{{{NomColonne}}}} pour personnaliser. Colonnes disponibles : {', '.join(['{{' + c + '}}' for c in colonnes_disponibles])}")
else:
    st.warning("⚠️ Importez d'abord un fichier Excel pour voir les colonnes disponibles")

sujet = st.text_input("📌 Objet de l'email", value="Bonjour {{Nom}}")

corps_template = st.text_area(
    "✏️ Corps du message (HTML supporté)",
    value="""<p>Bonjour {{Nom}},</p>
<p>Voici votre message personnalisé.</p>
<p>Cordialement,</p>""",
    height=200
)

# Aperçu avec la première ligne du fichier
if st.session_state.df is not None and len(st.session_state.df) > 0:
    with st.expander("👁️ Aperçu avec la première ligne de données"):
        apercu = corps_template
        premiere_ligne = st.session_state.df.iloc[0]
        for colonne in st.session_state.df.columns:
            apercu = apercu.replace(f"{{{{{colonne}}}}}", str(premiere_ligne[colonne]))

        if st.session_state.signature_bytes:
            apercu += "<br><br><em>[Signature image sera ajoutée ici]</em>"

        st.markdown(apercu, unsafe_allow_html=True)

st.divider()

# ============================================
# SECTION 5 : ENVOI D'UN EMAIL TEST
# ============================================
st.header("5️⃣ Envoyer un email de test")

col_test1, col_test2 = st.columns([3, 1])
with col_test1:
    email_test = st.text_input("📧 Adresse email pour le test")
with col_test2:
    st.write("")
    st.write("")
    bouton_test = st.button("🧪 Envoyer le test", use_container_width=True)

if bouton_test:
    if not sender_email or not password:
        st.error("❌ Veuillez remplir vos identifiants SMTP")
    elif not email_test:
        st.error("❌ Veuillez indiquer une adresse email de test")
    else:
        with st.spinner("Envoi en cours..."):
            # Remplacer les variables si un fichier existe, sinon texte brut
            corps_test = corps_template
            sujet_test = sujet
            if st.session_state.df is not None and len(st.session_state.df) > 0:
                premiere_ligne = st.session_state.df.iloc[0]
                for colonne in st.session_state.df.columns:
                    corps_test = corps_test.replace(f"{{{{{colonne}}}}}", str(premiere_ligne[colonne]))
                    sujet_test = sujet_test.replace(f"{{{{{colonne}}}}}", str(premiere_ligne[colonne]))

            success, message = envoyer_email(
                smtp_server=smtp_server,
                port=port,
                sender_email=sender_email,
                password=password,
                destinataire=email_test,
                sujet=sujet_test,
                corps_html=corps_test,
                signature_bytes=st.session_state.signature_bytes,
                use_ssl=use_ssl
            )

            if success:
                st.success(message)
            else:
                st.error(message)

st.divider()

# ============================================
# SECTION 6 : ENVOI GROUPÉ
# ============================================
st.header("6️⃣ Envoi groupé à tous les destinataires")

if st.session_state.df is not None:
    colonne_email = st.selectbox(
        "📧 Quelle colonne contient les adresses email ?",
        st.session_state.df.columns
    )

    nb_destinataires = len(st.session_state.df)
    st.info(f"📊 {nb_destinataires} destinataire(s) seront contactés")

    confirmation = st.checkbox("✅ Je confirme vouloir envoyer les emails à tous les destinataires")

    if st.button("🚀 Lancer l'envoi groupé", disabled=not confirmation, use_container_width=True):
        if not sender_email or not password:
            st.error("❌ Veuillez remplir vos identifiants SMTP")
        else:
            progress_bar = st.progress(0)
            status_text = st.empty()
            resultats = []

            for index, row in st.session_state.df.iterrows():
                # Personnalisation du sujet et du corps
                corps_perso = corps_template
                sujet_perso = sujet
                for colonne in st.session_state.df.columns:
                    valeur = str(row[colonne])
                    corps_perso = corps_perso.replace(f"{{{{{colonne}}}}}", valeur)
                    sujet_perso = sujet_perso.replace(f"{{{{{colonne}}}}}", valeur)

                destinataire = str(row[colonne_email])

                success, message = envoyer_email(
                    smtp_server=smtp_server,
                    port=port,
                    sender_email=sender_email,
                    password=password,
                    destinataire=destinataire,
                    sujet=sujet_perso,
                    corps_html=corps_perso,
                    signature_bytes=st.session_state.signature_bytes,
                    use_ssl=use_ssl
                )

                resultats.append({
                    "Destinataire": destinataire,
                    "Statut": "✅ Envoyé" if success else "❌ Échec",
                    "Détail": message
                })

                # Mise à jour de la barre de progression
                progression = (index + 1) / nb_destinataires
                progress_bar.progress(progression)
                status_text.text(f"Envoi en cours... {index + 1}/{nb_destinataires}")

                # Petite pause pour éviter de surcharger le serveur SMTP
                time.sleep(0.5)

            status_text.text("✅ Envoi terminé !")

            # Affichage des résultats
            st.subheader("📊 Résultats de l'envoi")
            df_resultats = pd.DataFrame(resultats)
            st.dataframe(df_resultats, use_container_width=True)

            nb_succes = len([r for r in resultats if "✅" in r["Statut"]])
            nb_echecs = len([r for r in resultats if "❌" in r["Statut"]])

            col_res1, col_res2 = st.columns(2)
            with col_res1:
                st.success(f"✅ {nb_succes} email(s) envoyé(s) avec succès")
            with col_res2:
                if nb_echecs > 0:
                    st.error(f"❌ {nb_echecs} échec(s)")
else:
    st.warning("⚠️ Veuillez d'abord importer un fichier Excel")

st.divider()
st.caption("💡 Astuce : testez toujours avec un email avant l'envoi groupé pour vérifier le rendu.")