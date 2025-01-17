import webview
import os
import ctypes
from generate_certificates import generate_certificates
from generate_agreements import generate_agreements
from utils import validate_file, logging_error, log_and_alert_error, set_app_icon

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

        broj_certifkata_ui = webview.windows[0].evaluate_js('document.getElementById("broj_certifkata").value')
        naziv_firme_ui = webview.windows[0].evaluate_js('document.getElementById("naziv_firme").value')
        adresa_firme_ui = webview.windows[0].evaluate_js('document.getElementById("adresa_firme").value')
        datum_dokumenta_ui = webview.windows[0].evaluate_js('document.getElementById("datum_dokumenta").value')
        grad_firme_ui = webview.windows[0].evaluate_js('document.getElementById("grad_firme").value')

        validate_file(excel_file_path[0], "Excel file")
        validate_file(word_template_path[0], "Word template")

        generate_certificates(
            excel_file_path[0], word_template_path[0], output_folder[0],
            broj_certifkata_ui, naziv_firme_ui, adresa_firme_ui, datum_dokumenta_ui, grad_firme_ui
        )
        return "Certificates generated successfully!"
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

        broj_certifkata = webview.windows[0].evaluate_js('document.getElementById("broj_certifkata").value')
        naziv_firme = webview.windows[0].evaluate_js('document.getElementById("naziv_firme").value')
        adresa_firme = webview.windows[0].evaluate_js('document.getElementById("adresa_firme").value')
        grad_firme = webview.windows[0].evaluate_js('document.getElementById("grad_firme").value')
        datum_dokumenta = webview.windows[0].evaluate_js('document.getElementById("datum_dokumenta").value')

        if not all([broj_certifkata, naziv_firme, adresa_firme, grad_firme, datum_dokumenta]):
            raise ValueError("Missing required UI input(s). Please ensure all fields are filled.")

        generate_agreements(
            excel_file_path[0],
            word_template_path[0],
            output_folder[0],
            broj_certifkata,
            naziv_firme,
            adresa_firme,
            grad_firme,
            datum_dokumenta
        )
        return "Agreements generated successfully!"
    except Exception as e:
        logging_error(f"Error in handle_agreements: {str(e)}")
        raise


class Api:
    def process_certificates(self):
        try:
            handle_certificates()
            webview.windows[0].evaluate_js('showResult("Certificates generated successfully!")')
        except Exception as e:
            log_and_alert_error(e)

    def process_agreements(self):
        try:
            handle_agreements()
            webview.windows[0].evaluate_js('showResult("Agreements generated successfully!")')
        except Exception as e:
            log_and_alert_error(e)

    def exit_application(self):
        logging_error("Application exited by user.", is_error=False)
        webview.windows[0].destroy()
        os._exit(0)

if __name__ == "__main__":
    try:
        set_app_icon("icon.ico")
        window = webview.create_window("Maherator", "index.html", width=800, height=600, js_api=Api())
        webview.start()
    except KeyboardInterrupt:
        logging_error("Application interrupted by user (Ctrl+C). Exiting...")
    except Exception as e:
        log_and_alert_error(e)
