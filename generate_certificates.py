from docx import Document
from datetime import datetime
import os
import pandas as pd


def replace_placeholder_in_paragraph(paragraph, placeholders):
    """
    Replaces placeholders in a paragraph, preserving formatting and handling split runs.
    """
    # Combine all runs' text in the paragraph
    full_text = "".join(run.text for run in paragraph.runs)

    # Replace placeholders in the combined text
    for placeholder, replacement in placeholders.items():
        full_text = full_text.replace(placeholder, replacement)

    # Clear all existing runs in the paragraph
    for run in paragraph.runs:
        run.text = ""

    # Add the updated text back to the paragraph
    paragraph.add_run(full_text)



def generate_certificates(excel_file_path, word_template_path, output_folder,
                          broj_certifkata_ui, naziv_firme_ui, adresa_firme_ui, datum_dokumenta_ui, grad_ui):
    # Load the Excel data
    data = pd.read_excel(excel_file_path)
    current_year = datetime.now().year

    # Load the Word template
    template_doc = Document(word_template_path)

    for index, row in data.iterrows():
        doc_copy = Document(word_template_path)
        ime_prezime= f"{row.get('ime')+("_")+row.get('prezime', '')}"
        # Prepare placeholders and their replacements
        placeholders = {
            "<<broj_certifkata>>": f"{broj_certifkata_ui}-{index + 1}/{current_year}",
            "<<datum_dokumenta>>": datum_dokumenta_ui,
            "<<naziv_firme>>": naziv_firme_ui,
            "<<adresa_firme>>": adresa_firme_ui,
            "<<ime_prezime>>": f"{row.get('ime', '').strip()} {row.get('prezime', '').strip()}",
            "<<datum_rodjenja>>": f"{row.get('datum_rodjenja')}",
            "<<adresa_stanovanja>>": f"{row.get('adresa_stanovanja')}",
            "<<grad>>": grad_ui,
            "<<vrsta_posla>>": row.get('vrsta_posla', '').strip()
        }

        # Replace placeholders in paragraphs
        for paragraph in doc_copy.paragraphs:
            replace_placeholder_in_paragraph(paragraph, placeholders)

        # Replace placeholders in tables (if any)
        for table in doc_copy.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        replace_placeholder_in_paragraph(paragraph, placeholders)

        # Save the updated document
        output_file = os.path.join(output_folder, f"Certifikat_{ime_prezime}.docx")
        doc_copy.save(output_file)
        print(f"Certifikati generisani: {output_file}")