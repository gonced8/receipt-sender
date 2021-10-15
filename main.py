import configparser
import datetime
from fpdf import FPDF
import locale
import os
import pygsheets
import smtplib
import ssl
import unidecode


def get_filename(name, date):
    name = unidecode.unidecode(name)
    name = name.lower()
    name = name.replace(" ", "")
    return name + "_" + date


def generate_receipt(n, name, address, nif, value, date, today):

    document = FPDF(orientation="L", format="A5")
    document.add_page()
    document.set_margins(20, 20, 20)
    document.set_auto_page_break(False)

    document.image("logo.png", x=5, y=10, w=50, type="PNG")

    document.ln()
    document.set_font("helvetica", "B", size=12)
    document.cell(
        w=0,
        h=6,
        txt="Sport Clube Badminton de Lisboa",
        align="C",
    )

    document.cell(w=0, h=6, txt=f"RECIBO Nº {n}", ln=1, align="R")
    document.ln()

    document.set_font("helvetica", size=12)
    document.multi_cell(
        w=0,
        h=5,
        txt="Parque de Jogos 1º de Maio\nAv. Rio de Janeiro\n1700-330 Lisboa\nE-mail: scbl@badmintonlisboa.pt\nContribuinte: XXXXXXXXX",
        align="C",
    )

    EURO = chr(128)
    long_date = datetime.datetime.strptime(date, "%Y-%m").strftime("%B de %Y")
    document.set_y(70)
    document.multi_cell(
        w=0,
        h=6,
        txt=f"RECEBI do Exmo(a). Sr(a). {name}, morada {address}, contribuinte nº {nif}, a quantia de {value}{EURO} referente à mensalidade de {long_date}.",
    )

    document.set_y(130)
    document.cell(w=0, h=6, txt=today)
    document.cell(w=0, h=6, txt="ASSINATURA", align="R")

    document.set_y(140)
    document.cell(w=0, h=0, txt="1/1", align="C")

    document.set_font("helvetica", size=8)
    document.set_y(110)
    document.cell(
        w=0,
        h=0,
        txt="(Este documento é meramente informativo e não serve de fatura.)",
        align="C",
    )

    filename = os.path.join("docs", get_filename(name, date) + ".pdf")
    document.output(filename)
    return filename


def send_email(config, name, receiver_email, nif, value, date, filename):
    smtp_server = config["EMAIL"]["server"]
    port = int(config["EMAIL"]["port"])  # For starttls
    sender_email = config["EMAIL"]["email"]
    password = config["EMAIL"]["password"]

    message = (
        f"From: {sender_email}\n"
        f"To: {receiver_email}\n"
        f"Subject: e-mail teste\n"
        "\n"
        "E aqui vai a mensagem."
    )

    # Create a secure SSL context
    context = ssl.create_default_context()

    with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, message)

    print(f"Sent receipt {filename} of {date} to {name} at {receiver_email}.")


def get_today_date():
    locale.setlocale(locale.LC_TIME, "pt_PT.UTF-8")
    week_days = [
        "domingo",
        "segunda-feira",
        "terça-feira",
        "quarta-feira",
        "quinta-feira",
        "sexta-feira",
        "sábado",
    ]
    today = datetime.date.today()
    week_day = int(today.strftime("%w"))
    week_day = week_days[week_day]
    today = week_day + today.strftime(", %d de %B de %Y")

    return today


def main():
    config = configparser.ConfigParser()
    config.read("config.ini")

    today = get_today_date()

    # Get sheet
    gc = pygsheets.authorize()
    sh = gc.open_by_key(config["GDRIVE"]["file_key"])

    # Get info page
    info = sh.worksheet("title", "Info")
    header = info.get_row(include_tailing_empty=False, row=1, returnas="matrix")
    n_info = len(header)

    # Get payments page
    payments = sh.worksheet("title", "Mensalidades")
    header = payments.get_row(include_tailing_empty=False, row=1, returnas="matrix")
    n_months = len(header)
    dates = header[1:]
    print("dates", dates)

    # Get aux variables page
    aux = sh.worksheet("title", "auxiliar")
    n = int(aux.get_value((1, 2)))

    # Initialize rows iterators
    info = iter(info)
    next(info)

    payments = iter(payments)
    next(payments)

    # Process each person
    for line, (person_info, person_payments) in enumerate(zip(info, payments), start=2):
        # Get person info
        name, email, address, nif, value = person_info[:n_info]
        print(line, name)
        # Get person payments
        check_name, *person_payments = person_payments[:n_months]

        # Check if inconsistency in pages
        assert (
            name == check_name
        ), f"Inconsistency in line {line}. Info has {name} but Mensalidades has {check_name}"

        # Filter already processed payments
        paid = [
            (i, date)
            for i, (date, value) in enumerate(zip(dates, person_payments))
            if value.upper() == "P"
        ]

        # Generate receipts
        filenames = []
        for i, date in paid:
            n += 1
            filename = generate_receipt(n, name, address, nif, value, date, today)
            filenames.append(filename)

        # Send e-mail and update payments page
        send_email(config, name, email, nif, value, date, filenames)
        wks.update_value((line, date_start + i), "E")

    # Update aux variables
    aux.update_value((1, 2), str(n))


if __name__ == "__main__":
    main()
