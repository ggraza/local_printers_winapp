"""
Printer handlers – receive pre-rendered HTML from the Frappe server,
convert to PDF via wkhtmltopdf, and silently print via SumatraPDF.
"""

import subprocess
import tempfile
import os
import logging
from logging.handlers import RotatingFileHandler

import pdfkit
import win32print

# ---------------------------------------------------------------------------
# Logging – file + console
# ---------------------------------------------------------------------------
APP_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_DIR = os.path.join(APP_DIR, "logs")
os.makedirs(LOG_DIR, exist_ok=True)

log = logging.getLogger("printer_handlers")
log.setLevel(logging.DEBUG)

_file_handler = RotatingFileHandler(
    os.path.join(LOG_DIR, "printer_handlers.log"),
    maxBytes=5 * 1024 * 1024,
    backupCount=5,
    encoding="utf-8",
)
_file_handler.setLevel(logging.DEBUG)
_file_handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
log.addHandler(_file_handler)


def get_local_printers() -> list[str]:
    """Return names of locally-installed printers."""
    return [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)]


def print_pdf_silent(pdf_path: str, printer_name: str, sumatra_pdf_path: str):
    """Print a PDF file silently using SumatraPDF."""
    command = (
        f'"{sumatra_pdf_path}" -print-to "{printer_name}" '
        f'-print-settings "noscale" "{pdf_path}"'
    )
    print(f"[PRINT] Sending PDF to printer '{printer_name}' ...")
    log.info("SumatraPDF command: %s", command)
    try:
        subprocess.run(command, shell=True, check=True)
        print(f"[PRINT] ✅ Sent '{pdf_path}' to printer '{printer_name}' successfully.")
        log.info("Sent %s to printer '%s'.", pdf_path, printer_name)
    except subprocess.CalledProcessError as exc:
        print(f"[PRINT] ❌ SumatraPDF FAILED for printer '{printer_name}': {exc}")
        log.error("SumatraPDF failed for '%s': %s", printer_name, exc)
    except Exception as exc:
        print(f"[PRINT] ❌ Unexpected error printing PDF: {exc}")
        log.error("Unexpected error printing PDF: %s", exc)


def html_to_pdf(html: str, wkhtmltopdf_path: str) -> str | None:
    """Convert an HTML string to a temporary PDF file. Returns the PDF path."""
    config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path)
    pdf_path = tempfile.mktemp(suffix=".pdf")
    options = {
        "no-outline": None,
        "encoding": "utf-8",
        "enable-local-file-access": None,
        "load-error-handling": "ignore",
        "load-media-error-handling": "ignore",
    }
    print(f"[PDF] Converting HTML to PDF ({len(html)} chars) ...")
    try:
        pdfkit.from_string(html, pdf_path, configuration=config, options=options)
        print(f"[PDF] ✅ Generated PDF: {pdf_path}")
        log.info("Generated PDF: %s", pdf_path)
        return pdf_path
    except Exception as exc:
        print(f"[PDF] ❌ wkhtmltopdf FAILED: {exc}")
        log.error("wkhtmltopdf failed: %s", exc)
        return None


def print_jobs(jobs: list[dict], config_data: dict) -> list[str]:
    """
    Process a list of print jobs received from the Frappe server.

    Each job dict contains:
      - html          : fully-rendered HTML (ready to print)
      - printer       : target printer system name
      - printer_ip    : (optional) network printer IP
      - invoice_name  : the Sales Invoice name
      - is_cashier    : whether this is the cashier copy
      - print_format  : the Print Format used server-side

    Returns a list of printer names that were printed to.
    """
    wkhtmltopdf_path = config_data["WKHTMLTOPDF"]
    sumatra_pdf_path = config_data.get(
        "SUMATRA_PDF_PATH", r"C:\Program Files\SumatraPDF\SumatraPDF.exe"
    )

    printed_to: list[str] = []

    if not isinstance(jobs, list):
        jobs = [jobs]

    print(f"\n{'='*60}")
    print(f"[JOBS] Processing {len(jobs)} print job(s)")
    print(f"{'='*60}")
    log.info("Processing %d print job(s)", len(jobs))

    for i, job in enumerate(jobs, 1):
        invoice_name = job.get("invoice_name", "unknown")
        printer_name = job.get("printer")
        print_format = job.get("print_format", "Standard")
        is_cashier = job.get("is_cashier", False)
        html = job.get("html")

        print(f"\n--- Job {i}/{len(jobs)} ---")
        print(f"  Invoice   : {invoice_name}")
        print(f"  Printer   : {printer_name}")
        print(f"  Format    : {print_format}")
        print(f"  Is Cashier: {is_cashier}")
        print(f"  HTML      : {'Yes (' + str(len(html)) + ' chars)' if html else 'NO ❌'}")

        log.info(
            "Job %d/%d – invoice=%s printer=%s format=%s is_cashier=%s html_len=%s",
            i, len(jobs), invoice_name, printer_name, print_format, is_cashier,
            len(html) if html else 0,
        )

        if not html:
            print(f"  ⚠️  SKIPPED – no HTML content")
            log.warning("Job for invoice %s has no HTML, skipping.", invoice_name)
            continue

        if not printer_name:
            print(f"  ⚠️  SKIPPED – no printer name")
            log.warning("Job for invoice %s has no printer, skipping.", invoice_name)
            continue

        pdf_path = html_to_pdf(html, wkhtmltopdf_path)
        if pdf_path:
            print_pdf_silent(pdf_path, printer_name, sumatra_pdf_path)
            printed_to.append(printer_name)

            # Clean up temp PDF
            try:
                os.remove(pdf_path)
                log.info("Cleaned up temp PDF: %s", pdf_path)
            except OSError:
                pass
        else:
            print(f"  ❌ PDF generation failed – nothing sent to printer")

    print(f"\n{'='*60}")
    print(f"[JOBS] Done. Printed to: {printed_to if printed_to else 'NONE'}")
    print(f"{'='*60}\n")
    log.info("Finished processing. Printed to: %s", printed_to)

    return printed_to
