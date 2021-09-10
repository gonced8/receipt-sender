import configparser
import datetime
from fpdf import FPDF
import locale
import os
import pygsheets
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

    document.set_y(120)
    document.cell(w=0, h=6, txt=today)
    document.cell(w=0, h=6, txt="ASSINATURA", align="R")

    document.set_y(140)
    document.cell(w=0, h=0, txt="1/1", align="C")

    filename = os.path.join("docs", get_filename(name, date) + ".pdf")
    document.output(filename)
    return filename


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

config = configparser.ConfigParser()
config.read("config.ini")

gc = pygsheets.authorize()

sh = gc.open_by_key(config["DEFAULT"]["file_key"])
wks = sh.sheet1

header = wks.get_row(include_tailing_empty=False, row=1, returnas="matrix")
ncols = len(header)

date_start = 5
dates = header[date_start - 1 :]
print("dates", dates)

n = int(sh.worksheet(value=1).get_value((1, 2)))

rows = iter(wks)
next(rows)
for line, row in enumerate(rows, start=2):
    name, email, address, nif, *payments = row[:ncols]
    print(line, name, email, payments)

    paid = [
        (i, date)
        for i, (date, value) in enumerate(zip(dates, payments))
        if value.upper() == "X"
    ]

    for i, date in paid:
        n += 1
        value = 15
        filename = generate_receipt(n, name, address, nif, value, date, today)
        print(f"Sent receipt {filename} of {date} to {name} at {email}.")
        # wks.update_value((line, date_start + i), "#")

sh.worksheet(value=1).update_value((1, 2), str(n))
