from babel.dates import format_date
import configparser
import datetime
from email.message import EmailMessage
from fpdf import FPDF
import os
from pathlib import Path
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

    document.image("assets/logo.png", x=5, y=10, w=50, type="PNG")

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
        txt="Parque de Jogos 1º de Maio\nAv. Rio de Janeiro\n1700-330 Lisboa\nE-mail: scbl@badmintonlisboa.pt\nContribuinte: 516248758",
        align="C",
    )

    EURO = chr(128)
    long_date = format_date(
        datetime.datetime.strptime(date, "%Y-%m"), "MMMM 'de' yyyy", locale="pt_PT"
    )
    document.set_y(70)
    document.multi_cell(
        w=0,
        h=6,
        txt=f"RECEBI do Exmo(a). Sr(a). {name}, morada {address}, contribuinte nº {nif}, a quantia de {value}{EURO} referente à mensalidade de {long_date}.",
    )

    document.set_y(125)
    document.cell(w=0, h=6, txt=today)
    # document.cell(w=0, h=6, txt="ASSINATURA", align="R")
    document.image("assets/stamp.png", x=140, y=120, w=50, type="PNG")

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
    Path("docs").mkdir(exist_ok=True)
    document.output(filename)
    return filename


def send_email(config, password, name, receiver_email, nif, value, filenames):
    smtp_server = config["server"]
    port = int(config["port"])  # For starttls
    sender_email = config["email"]
    # password = config["password"]
    sender = config["sender"]
    signature_filename = config["signature"]
    template_filename = config["template"]

    msg = EmailMessage()

    # Header
    msg["Subject"] = "Recibos mensalidade SCBL"
    msg["From"] = f"{sender} <{sender_email}>"
    msg["To"] = f"{name} <{receiver_email}>"

    # Plain text Body
    receipts = [os.path.split(filename)[-1] for filename in filenames]

    receipts_list = "- " + "\n- ".join(receipts)

    message = (
        "Segue em anexo os recibos:\n"
        f"{receipts_list}\n"
        "\n"
        "SCBL\n"
        "\n"
        "(Este e-mail foi gerado automaticamente.)"
    )
    msg.set_content(message)

    # HTML Body
    """
    with open(signature_filename, "r") as f:
        signature = f.read()

    with open(template_filename, "r") as f:
        template = f.read()

    html_message = template.format(
        content=message.replace("\n", "<br>"), signature=signature
    )
    msg.add_alternative(html_message, subtype="html")
    """

    # Attachments
    for filename, receipt in zip(filenames, receipts):
        with open(filename, "rb") as f:
            msg.add_attachment(
                f.read(),
                maintype="application",
                subtype="octate-stream",
                filename=receipt,
            )

    # msg = f"Subject: teste\nFrom: {sender_email}\nTo: {receiver_email}\n\nOla"

    # Create a secure SSL context
    context = ssl.create_default_context()

    # Send e-mail
    with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
        server.login(sender_email, password)
        # server.set_debuglevel(1)
        server.send_message(msg)

    print(f"Sent receipt {' '.join(filenames)} to {name} at {receiver_email}.")


def get_today_date():
    today = datetime.date.today()
    today = format_date(today, "full", locale="pt_PT")
    return today


def main():
    config = configparser.ConfigParser()
    config.read("config.ini")

    today = get_today_date()

    # Get sheet
    gc = pygsheets.authorize()
    sh = gc.open_by_key(config["GDRIVE"]["file_key"])

    # Get info page
    info_sheet = sh.worksheet("title", "Info")
    header = info_sheet.get_row(include_tailing_empty=False, row=1, returnas="matrix")
    n_info = len(header)

    # Get payments page
    payments_sheet = sh.worksheet("title", "Mensalidades")
    header = payments_sheet.get_row(
        include_tailing_empty=False, row=1, returnas="matrix"
    )
    n_months = len(header)
    dates = header[1:]
    print("dates", dates)

    # Get aux variables page
    aux_sheet = sh.worksheet("title", "auxiliar")
    n = int(aux_sheet.get_value((1, 2)))

    # Initialize rows iterators
    info = iter(info_sheet)
    next(info)

    payments = iter(payments_sheet)
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
            for i, (date, state) in enumerate(zip(dates, person_payments), start=2)
            if state.upper() == "P"
        ]

        # Generate receipts
        filenames = []
        for _, date in paid:
            n += 1
            filename = generate_receipt(n, name, address, nif, value, date, today)
            filenames.append(filename)

        # Send e-mail
        if filenames:
            send_email(config, name, email, nif, value, filenames)

        # Update payments
        for i, _ in paid:
            payments_sheet.update_value((line, i), "E")

    # Update aux variables
    aux_sheet.update_value((1, 2), str(n))


if __name__ == "__main__":
    main()
