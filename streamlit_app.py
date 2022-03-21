from email.message import EmailMessage
import io
import openpyxl
import os
import pandas as pd
from pathlib import Path
import smtplib
import ssl
import streamlit as st
import urllib.request

from utils import *

# Download images
# if not Path("assets").is_dir():
#    Path("assets").mkdir(exist_ok=True)

# Logo
if not Path("assets/logo.png").is_file():
    urllib.request.urlretrieve(st.secrets["IMAGES"]["logo"], "assets/logo.png")
    print("Downloaded logo")

# Stamp
if not Path("assets/stamp.png").is_file():
    urllib.request.urlretrieve(st.secrets["IMAGES"]["stamp"], "assets/stamp.png")
    print("Downloaded stamp")

st.title("Gerador de Recibos")

password = st.text_input("Enter a password", type="password")

if password == st.secrets["OTHER"]["authentication_password"]:
#if True:
    uploaded_file = st.file_uploader(
        "Selecionar ficheiro",
        "xlsx",
    )

    if uploaded_file is not None and st.button("Run"):
        data = pd.read_excel(uploaded_file, sheet_name=None)
        for sheet in data:
            data[sheet].fillna("", inplace=True)

        today = get_today_date()

        # Process each person
        for index, receipt in data["Recibos"].iterrows():
            print(
                index,
                receipt["Nº de Sócio"],
                receipt["Nome do Atleta"],
                receipt["E-mail"],
                receipt["Contribuinte"],
                receipt["Data de recebimento"],
                receipt["Valor"],
                receipt["Descritivo"],
                receipt["Nº do Recibo"],
                receipt["Status"],
            )

            # Skip already processed receipts
            if receipt["Status"].upper() != "P":
                continue

            # Generate receipt
            filename = generate_receipt(
                receipt["Data de recebimento"],
                receipt["Valor"],
                receipt["Descritivo"],
                receipt["Nome do Atleta"],
                receipt["Nº de Sócio"],
                receipt["Contribuinte"],
                receipt["Nº do Recibo"],
                today,
            )

            # Send e-mail
            if receipt["E-mail"]:
                send_email(
                    st.secrets["EMAIL"],
                    receipt["Nome do Atleta"],
                    receipt["E-mail"],
                    receipt["Valor"],
                    receipt["Descritivo"],
                    filename,
                )
            else:
                print(f"{receipt['Nome do Atleta']} has not an e-mail...")

            # Update payments
            data["Recibos"].at[index, "Status"] = "E"

        # Get output filename
        input_name, input_extension = os.path.splitext(uploaded_file.name)
        output_name = input_name + "_atualizado" + input_extension

        # Generate new excel
        output = io.BytesIO()
        writer = pd.ExcelWriter(output, engine="openpyxl")
        data["Recibos"].to_excel(writer, "Recibos", index=False)
        writer.save()
        xlsx_data = output.getvalue()

        # Download updated file
        tmp_download_link = download_link(
            xlsx_data, output_name, "Click here to download"
        )
        st.markdown(tmp_download_link, unsafe_allow_html=True)

elif password:
    st.text("Password errada")
