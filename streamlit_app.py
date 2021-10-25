from email.message import EmailMessage
import smtplib
import ssl
import streamlit as st

st.title("Gerador de Recibos")

if st.button("Send e-mail"):
    smtp_server = st.secrets["EMAIL"]["server"]
    port = int(st.secrets["EMAIL"]["port"])  # For starttls
    sender_email = st.secrets["EMAIL"]["email"]
    password = st.secrets["EMAIL"]["password"]
    sender = st.secrets["EMAIL"]["sender"]

    msg = EmailMessage()

    # Header
    msg["Subject"] = "Recibos mensalidade SCBL"
    msg["From"] = f"{sender} <{sender_email}>"
    msg["To"] = ""

    # Plain text Body
    message = (
        "Segue em anexo os recibos:\n"
        f"TESTE\n"
        "\n"
        "SCBL\n"
        "\n"
        "(Este e-mail foi gerado automaticamente.)"
    )
    msg.set_content(message)

    # Create a secure SSL context
    context = ssl.create_default_context()

    # Send e-mail
    with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
        server.login(sender_email, password)
        # server.set_debuglevel(1)
        server.send_message(msg)

    print(f"Sent receipt {' '.join(filenames)} to {name} at {receiver_email}.")
