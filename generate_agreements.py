from docx import Document
from datetime import datetime
import pandas as pd
import win32com.client as win32
import logging

# Configure logging
logging.basicConfig(
    filename="application.log",  # Log file
    level=logging.INFO,  # Log level
    format="%(asctime)s - %(levelname)s - %(message)s",  # Log format
)

def load_word_template(file_path):
    """
    Loads a Word template, converting .doc to .docx if necessary.
    """
    if file_path.endswith(".doc"):
        logging.info(f"Converting .doc file: {file_path}")
        file_path = convert_doc_to_docx(file_path)  # Convert to .docx
        if not file_path:
            logging.error("Conversion from .doc to .docx failed.")
            raise ValueError("Conversion from .doc to .docx failed.")
    return Document(file_path)


def convert_doc_to_docx(doc_file_path):
    """
    Converts a .doc file to .docx format.
    """
    word = win32.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(doc_file_path)
        docx_file_path = doc_file_path + "x"  # Append 'x' to save as .docx
        doc.SaveAs2(docx_file_path, FileFormat=16)  # 16 is the file format for .docx
        doc.Close()
        logging.info(f"Converted {doc_file_path} to {docx_file_path}")
        return docx_file_path
    except Exception as e:
        logging.error(f"Error converting {doc_file_path} to .docx: {e}")
        return None
    finally:
        word.Quit()


def replace_placeholder(paragraph, placeholders):
    """
    Replaces placeholders in a paragraph while preserving its formatting.
    """
    for run in paragraph.runs:
        for placeholder, replacement in placeholders.items():
            if placeholder in run.text:
                run.text = run.text.replace(placeholder, replacement)


def replace_placeholders_in_table_row(row, placeholders):
    """
    Replaces placeholders in a specific table row while preserving formatting.
    """
    for cell in row.cells:
        for paragraph in cell.paragraphs:
            replace_placeholder(paragraph, placeholders)


def populate_first_table(table, data, static_placeholders):
    """
    Populates the first table with dynamic rows for each entry in the Excel data.
    """
    template_row = table.rows[1]

    for index, row_data in data.iterrows():
        placeholders = {
            "<<redni_broj>>": str(index + 1),
            "<<ime>>": row_data.get("ime", "").strip(),
            "<<ocevo_ime>>": row_data.get("ocevo_ime", "").strip(),
            "<<prezime>>": row_data.get("prezime", "").strip(),
            "<<datum_rodjenja>>": row_data.get("datum_rodjenja", "").strip(),
            "<<naziv_firme>>": static_placeholders["<<naziv_firme>>"],
            "<<adresa_firme>>": static_placeholders["<<adresa_firme>>"],
            "<<grad>>": static_placeholders["<<grad>>"],
        }

        new_row = table.add_row()
        for i, cell in enumerate(new_row.cells):
            cell.text = template_row.cells[i].text  # Copy the structure

        replace_placeholders_in_table_row(new_row, placeholders)

    table._tbl.remove(template_row._element)


def populate_second_table(table, data, static_placeholders):
    """
    Populates the second table, with `<<vrsta_posla>>` written only once and other placeholders written per row.
    """
    template_row_1 = table.rows[0]
    template_row_2 = table.rows[1]

    common_placeholder = data.iloc[0].get("vrsta_posla", "").strip()
    first_row = table.add_row().cells[0]
    first_row.text = f"BEZBJEDAN I SIGURAN RAD NA RADNOM MJESTU {common_placeholder}"

    for index, row_data in data.iterrows():
        placeholders = {
            "<<broj_certifkata>>": static_placeholders["<<broj_certifkata>>"],
            "<<redni_broj>>": str(index + 1),
            "<<godina>>": static_placeholders["<<godina>>"],
        }
        second_row = table.add_row().cells[0]
        second_row.text = f"te se izdaje UVJERENJE br. {placeholders['<<broj_certifkata>>']}-{placeholders['<<redni_broj>>']}/{placeholders['<<godina>>']}"

    table._tbl.remove(template_row_1._element)
    table._tbl.remove(template_row_2._element)


def generate_agreements(excel_file_path, word_template_path, output_file_path,
                        broj_certifkata, naziv_firme, adresa_firme, grad, datum_dokumenta):
    """
    Generates agreements based on a Word template, Excel data, and UI inputs.
    """
    try:
        data = pd.read_excel(excel_file_path)
        current_year = datetime.now().year

        doc = load_word_template(word_template_path)
        broj_dokumenta = f"{broj_certifkata}/{current_year}"

        static_placeholders = {
            "<<broj_dokumenta>>": broj_dokumenta,
            "<<broj_certifkata>>": broj_certifkata,
            "<<naziv_firme>>": naziv_firme,
            "<<adresa_firme>>": adresa_firme,
            "<<grad>>": grad,
            "<<datum_dokumenta>>": datum_dokumenta,
            "<<godina>>": str(current_year),
        }

        for paragraph in doc.paragraphs:
            replace_placeholder(paragraph, static_placeholders)

        first_table = next((tbl for tbl in doc.tables if "<<redni_broj>>" in tbl.rows[1].cells[0].text), None)
        if first_table:
            logging.info("First table found. Populating...")
            populate_first_table(first_table, data, static_placeholders)
        else:
            logging.warning("First table not found.")

        second_table = next((tbl for tbl in doc.tables if any("<<vrsta_posla>>" in cell.text for row in tbl.rows for cell in row.cells)), None)
        if second_table:
            logging.info("Second table found. Populating...")
            populate_second_table(second_table, data, static_placeholders)
        else:
            logging.warning("Second table not found.")

        doc.save(output_file_path)
        logging.info(f"Agreement document saved: {output_file_path}")
        print(f"Zapisnik generisan: {output_file_path}")

    except PermissionError:
        logging.error(f"Permission denied. Ensure the file '{output_file_path}' is not open.")
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
