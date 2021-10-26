from email.message import EmailMessage
import openpyxl
import pandas as pd
from pathlib import Path
import requests
import smtplib
import ssl
import streamlit as st

from utils import *

# Download images
# if not Path("assets").is_dir():
#    Path("assets").mkdir(exist_ok=True)

# Logo
if not Path("assets/logo.png").is_file():
    img_data = requests.get(st.secrets["IMAGES"]["logo"]).content
    with open("assets/logo.png", "wb") as handler:
        handler.write(img_data)

# Stamp
if not Path("assets/stamp.png").is_file():
    img_data = requests.get(st.secrets["IMAGES"]["stamp"]).content
    with open("assets/stamp.png", "wb") as handler:
        handler.write(img_data)

st.title("Gerador de Recibos")

password = st.text_input("Enter a password", type="password")

uploaded_file = st.file_uploader("Selecionar ficheiro")
if uploaded_file is not None and st.button("RUN"):
    data = pd.read_excel(uploaded_file, sheet_name=None)

    today = get_today_date()

    # Get dates
    dates = data["Mensalidades"].columns[1:].values
    n_months = len(dates)
    print(f"{dates=}")

    # Get aux variables
    data["auxiliar"] = data["auxiliar"].set_index("Descrição")
    n = data["auxiliar"].loc["Contador"].item()

    # Process each person
    for (index, person_info), (_, person_payments) in zip(
        data["Info"].iterrows(),
        data["Mensalidades"].iterrows(),
    ):
        # Get person info
        print(
            index,
            person_info["Nome"],
            person_info["E-mail"],
            person_info["Contribuinte"],
            person_info["Mensalidade"],
        )

        # Get person payments
        print(person_payments[1:])

        # Check if inconsistency in pages
        assert (
            person_info["Nome"] == person_payments["Nome"]
        ), f"Inconsistency in line {index+1}. Info has {person_info['Nome']} but Mensalidades has {person_payments['Nome']}"

        # Get payments to send e-mail
        payments = person_payments[1:][person_payments[1:] == "P"].keys()
        if payments.empty:
            continue

        # Generate receipts
        filenames = []
        print(payments)
        for date in payments:
            print(date)
            n += 1
            filename = generate_receipt(
                n,
                person_info["Nome"],
                person_info["Morada"],
                person_info["Contribuinte"],
                person_info["Mensalidade"],
                date,
                today,
            )
            filenames.append(filename)

        # Send e-mail
        if filenames:
            send_email(
                st.secrets["EMAIL"],
                password,
                person_info["Nome"],
                person_info["E-mail"],
                person_info["Contribuinte"],
                person_info["Mensalidade"],
                filenames,
            )

        # Update payments
        person_payments[payments] = "E"
