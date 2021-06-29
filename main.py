import configparser
from fpdf import FPDF
import os
import pygsheets
import unidecode


def get_filename(name, date):
    name = unidecode.unidecode(name)
    name = name.lower()
    name = name.replace(" ", "")
    return name + "_" + date


def generate_receipt(name, email, date):
    document = FPDF()
    document.add_page()
    document.set_font("helvetica", size=12)
    document.cell(txt=f"{name} pagou a mensalidade de {date}. Enviado para {email}.")
    filename = os.path.join("docs", get_filename(name, date) + ".pdf")
    document.output(filename)
    return filename


config = configparser.ConfigParser()
config.read("config.ini")

gc = pygsheets.authorize()

sh = gc.open_by_key(config["DEFAULT"]["file_key"])
wks = sh.sheet1

header = wks.get_row(include_tailing_empty=False, row=1, returnas="matrix")
ncols = len(header)

date_start = 3
dates = header[date_start - 1 :]
print("dates", dates)

rows = iter(wks)
next(rows)
for line, row in enumerate(rows, start=2):
    name, email, *payments = row[:ncols]
    print(line, name, email, payments)

    paid = [
        (i, date)
        for i, (date, value) in enumerate(zip(dates, payments))
        if value.upper() == "X"
    ]

    for i, date in paid:
        filename = generate_receipt(name, email, date)
        print(f"Sent receipt {filename} of {date} to {name} at {email}.")
        wks.update_value((line, date_start + i), "#")
