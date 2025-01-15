import webview
import os
import ctypes
from generate_certificates import generate_certificates
from generate_agreements import generate_agreements
import win32com.client as win32
from docx import Document


def convert_doc_to_docx(doc_file_path):
    """
    Converts a .doc file to .docx format using Microsoft Word.
    """
    word = win32.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(doc_file_path)
        docx_file_path = doc_file_path + "x"  # Add 'x' to save as .docx
        doc.SaveAs2(docx_file_path, FileFormat=16)  # FileFormat 16 is for .docx
        doc.Close()
        print(f"Converted {doc_file_path} to {docx_file_path}")
        return docx_file_path
    except Exception as e:
        print(f"Error converting {doc_file_path} to .docx: {e}")
        return None
    finally:
        word.Quit()

def load_word_template(file_path):
    """
    Loads a Word template, converting .doc to .docx if necessary.
    """
    if file_path.endswith(".doc"):
        print(f"Converting .doc file: {file_path}")
        file_path = convert_doc_to_docx(file_path)  # Convert to .docx
        if not file_path:
            raise ValueError("Conversion from .doc to .docx failed.")
    return Document(file_path)



# Function to handle certificates
def handle_certificates():
    excel_file_path = webview.windows[0].create_file_dialog(webview.OPEN_DIALOG, file_types=['Excel files (*.xlsx)'])
    if not excel_file_path:
        print("No Excel file selected.")
        return

    word_template_path = webview.windows[0].create_file_dialog(webview.OPEN_DIALOG, file_types=['Word files (*.docx;*.doc)'])
    if not word_template_path:
        print("No Word template selected.")
        return

    output_folder = webview.windows[0].create_file_dialog(webview.FOLDER_DIALOG)
    if not output_folder:
        print("No output folder selected.")
        return

    broj_certifkata_ui = webview.windows[0].evaluate_js('document.getElementById("broj_certifkata").value')
    naziv_firme_ui = webview.windows[0].evaluate_js('document.getElementById("naziv_firme").value')
    adresa_firme_ui = webview.windows[0].evaluate_js('document.getElementById("adresa_firme").value')
    datum_dokumenta_ui = webview.windows[0].evaluate_js('document.getElementById("datum_dokumenta").value')
    grad_ui = webview.windows[0].evaluate_js('document.getElementById("grad").value')

    generate_certificates(
        excel_file_path[0], word_template_path[0], output_folder[0],
        broj_certifkata_ui, naziv_firme_ui, adresa_firme_ui, datum_dokumenta_ui, grad_ui
    )


def handle_agreements():
    # File selection dialogs
    excel_file_path = webview.windows[0].create_file_dialog(webview.OPEN_DIALOG, file_types=['Excel files (*.xlsx)'])
    if not excel_file_path:
        print("No Excel file selected.")
        return

    word_template_path = webview.windows[0].create_file_dialog(webview.OPEN_DIALOG, file_types=['Word files (*.docx)'])
    if not word_template_path:
        print("No Word template selected.")
        return

    output_folder = webview.windows[0].create_file_dialog(webview.FOLDER_DIALOG)
    if not output_folder:
        print("No output folder selected.")
        return

    # Retrieve UI inputs
    
    broj_certifkata = webview.windows[0].evaluate_js('document.getElementById("broj_certifkata").value')
    naziv_firme = webview.windows[0].evaluate_js('document.getElementById("naziv_firme").value')
    adresa_firme = webview.windows[0].evaluate_js('document.getElementById("adresa_firme").value')
    grad = webview.windows[0].evaluate_js('document.getElementById("grad").value')
    datum_dokumenta = webview.windows[0].evaluate_js('document.getElementById("datum_dokumenta").value')

    # Ensure all required UI inputs are provided
    if not all([ broj_certifkata, naziv_firme, adresa_firme, grad, datum_dokumenta]):
        print("Missing required UI input(s). Please ensure all fields are filled.")
        return

    # Generate agreements
    output_file_path = os.path.join(output_folder[0], "Generated_Agreements.docx")
    generate_agreements(
        excel_file_path[0],
        word_template_path[0],
        output_file_path,
        broj_certifkata,
        naziv_firme,
        adresa_firme,
        grad,
        datum_dokumenta
    )



# API for JavaScript
class Api:
    def process_certificates(self):
        handle_certificates()

    def process_agreements(self):
        handle_agreements()


# Main PyWebview setup
if __name__ == "__main__":
    # Path to your .ico file
    app_icon_path = os.path.abspath("icon.ico")  # Ensure the path to your .ico file is correct

    # Use ctypes to set the app icon on Windows
    if os.name == "nt":  # Check if the OS is Windows
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Maherator")
        ctypes.windll.user32.LoadIconW(0, app_icon_path)

    # Create the PyWebView window
    window = webview.create_window(
        "Maherator",  # Window title
        "index.html",  # Path to your HTML file
        width=800,
        height=600,
    )

    webview.start()