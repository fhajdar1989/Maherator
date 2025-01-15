import webview
import os
import ctypes
from generate_certificates import generate_certificates
from generate_agreements import generate_agreements
from docx import Document
import win32com.client as win32
import traceback

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
        logging_error(f"Converted {doc_file_path} to {docx_file_path}")
        return docx_file_path
    except Exception as e:
        logging_error(f"Error converting {doc_file_path} to .docx: {e}")
        return None
    finally:
        word.Quit()
        del word  # Ensure the COM object is released


def load_word_template(file_path):
    """
    Loads a Word template, converting .doc to .docx if necessary.
    """
    if file_path.endswith(".doc"):
        logging_error(f"Converting .doc file: {file_path}")
        file_path = convert_doc_to_docx(file_path)
        if not file_path:
            raise ValueError("Conversion from .doc to .docx failed.")
    return Document(file_path)

def handle_certificates():
    try:
        excel_file_path = webview.windows[0].create_file_dialog(webview.OPEN_DIALOG, file_types=['Excel files (*.xlsx)'])
        if not excel_file_path:
            raise ValueError("No Excel file selected.")

        word_template_path = webview.windows[0].create_file_dialog(webview.OPEN_DIALOG, file_types=['Word files (*.docx;*.doc)'])
        if not word_template_path:
            raise ValueError("No Word template selected.")

        output_folder = webview.windows[0].create_file_dialog(webview.FOLDER_DIALOG)
        if not output_folder:
            raise ValueError("No output folder selected.")

        # Retrieve inputs from the UI
        broj_certifkata_ui = webview.windows[0].evaluate_js('document.getElementById("broj_certifkata").value')
        naziv_firme_ui = webview.windows[0].evaluate_js('document.getElementById("naziv_firme").value')
        adresa_firme_ui = webview.windows[0].evaluate_js('document.getElementById("adresa_firme").value')
        datum_dokumenta_ui = webview.windows[0].evaluate_js('document.getElementById("datum_dokumenta").value')
        grad_ui = webview.windows[0].evaluate_js('document.getElementById("grad").value')

        # Generate certificates
        generate_certificates(
            excel_file_path[0], word_template_path[0], output_folder[0],
            broj_certifkata_ui, naziv_firme_ui, adresa_firme_ui, datum_dokumenta_ui, grad_ui
        )
        return "Certifikati su uspješno generisani!"
    except Exception as e:
        logging_error(f"Error in handle_certificates: {str(e)}")
        raise ValueError(f"Failed to generate certificates: {str(e)}")

def handle_agreements():
    try:
        excel_file_path = webview.windows[0].create_file_dialog(webview.OPEN_DIALOG, file_types=['Excel files (*.xlsx)'])
        if not excel_file_path:
            raise ValueError("No Excel file selected.")

        word_template_path = webview.windows[0].create_file_dialog(webview.OPEN_DIALOG, file_types=['Word files (*.docx)'])
        if not word_template_path:
            raise ValueError("No Word template selected.")

        output_folder = webview.windows[0].create_file_dialog(webview.FOLDER_DIALOG)
        if not output_folder:
            raise ValueError("No output folder selected.")

        # Retrieve inputs from the UI
        broj_certifkata = webview.windows[0].evaluate_js('document.getElementById("broj_certifkata").value')
        naziv_firme = webview.windows[0].evaluate_js('document.getElementById("naziv_firme").value')
        adresa_firme = webview.windows[0].evaluate_js('document.getElementById("adresa_firme").value')
        grad = webview.windows[0].evaluate_js('document.getElementById("grad").value')
        datum_dokumenta = webview.windows[0].evaluate_js('document.getElementById("datum_dokumenta").value')

        # Validate inputs
        if not all([broj_certifkata, naziv_firme, adresa_firme, grad, datum_dokumenta]):
            raise ValueError("Missing required UI input(s). Please ensure all fields are filled.")

        # Generate agreements
        output_file_path = os.path.join(output_folder[0], f"Zapisnik_{datum_dokumenta.replace('.', '_')}.docx")
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
        return "Zapisnik je uspješno kreiran!"
    except Exception as e:
        logging_error(f"Error in handle_agreements: {str(e)}")
        raise ValueError(f"Failed to generate agreements: {str(e)}")

class Api:
    def process_certificates(self):
        try:
            result = handle_certificates()
            webview.windows[0].evaluate_js(f'showResult("{result}")')
        except Exception as e:
            log_and_alert_error(e)

    def process_agreements(self):
        try:
            result = handle_agreements()
            webview.windows[0].evaluate_js(f'showResult("{result}")')
        except Exception as e:
            log_and_alert_error(e)

    def exit_application(self):
        print("Exiting application...")
        logging_error("Application exited by user.", is_error=False)
        webview.windows[0].destroy()
        os._exit(0)


def logging_error(message, is_error=True):
    """
    Logs an error message to both the console and a log file if it's an error.
    Normal application behaviors (like exits) won't be logged.
    
    Args:
        message (str): The message to log.
        is_error (bool): Whether the message is an error. Defaults to True.
    """
    if is_error:
        print(message)
        with open("error.log", "a") as log_file:
            log_file.write(f"{message}\n")


def log_and_alert_error(e):
    error_message = f"Error: {str(e)}"
    logging_error(error_message)
    try:
        webview.windows[0].evaluate_js(f'showResult("{error_message}")')
    except webview.errors.JavascriptException:
        print(f"Failed to call showResult in JavaScript. Error: {error_message}")



def set_app_icon(icon_path):
    """
    Set the application icon for the taskbar (Windows-specific).
    """
    if os.name == "nt":
        myappid = 'Maherator.App'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

if __name__ == "__main__":
    try:
        set_app_icon("icon.ico")
        window = webview.create_window("Maherator", "index.html", width=800, height=600, js_api=Api())
        webview.start()
    except Exception as e:
        log_and_alert_error(e)
