#!/usr/bin/env python

import csv
from datetime import datetime
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, NumberObject, TextStringObject

def alternate_fill(page, fields):
    # The normal function breaks on certain text fields for some reason
    for i in range(len(page["/Annots"])):
        writer_annot = page["/Annots"][i].get_object()
        for field in fields:
            if writer_annot.get("/T") == field:
                writer_annot.update({
                    NameObject("/V"): TextStringObject(fields[field])
                })

def multiline(string):
    return string.replace("\\n", "\r")

def parse_csv_row(row, out_file):
    reader = PdfReader("hh-Cheque-Request(A).pdf")
    writer = PdfWriter()

    page = reader.pages[0]
    fields = reader.get_fields()
    writer.append(reader)
    writer.update_page_form_field_values(writer.pages[0], {
        "payable": row[0],
        "ADDRESS": multiline(row[1]),
        "date": datetime.today().strftime("%m/%d/%y"),
        "PURPOSE": multiline(row[3])
    })
    alternate_fill(writer.pages[0], {
        "CE 1": "836420 - Gifts/Goodwill/Prizes",
        "CC1": "H5520",
        "ACTIVITYASSIGNMENTRow1": multiline(row[2])
    })

    print(f"Writing '{out_file}'...")
    with open(out_file, "wb") as out:
        writer.write(out)

if __name__ == "__main__":
    print("Opening 'data.csv'...")
    with open("data.csv", newline="") as data:
        csv_reader = csv.reader(data)
        date = datetime.today().strftime("%Y-%m-%d")
        for i, row in enumerate(csv_reader):
            parse_csv_row(row, f"cheque_request_{date}_{i + 1}.pdf")
