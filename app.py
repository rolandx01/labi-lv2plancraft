"""
app.py - Flask Webapp für LV → Plancraft Konverter

Production-kompatibel: läuft mit gunicorn, lokal mit python app.py.
Statische Uploads/Downloads landen in /tmp (Render-Read-Only-Dateisystem).
"""

import os
import uuid
import logging
from pathlib import Path
from datetime import datetime

from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename

from lv_parser import parse_pdf
from xlsx_writer import schreibe_xlsx

# Logging konfigurieren (Render sammelt stdout/stderr)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "labi-lv2plancraft-dev-key-change-in-prod")
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB max

# Upload/Download-Verzeichnisse
# Wichtig: /tmp ist auf Render der einzige beschreibbare Ort
UPLOAD_DIR = Path(os.environ.get("UPLOAD_DIR", "/tmp/labi-uploads"))
DOWNLOAD_DIR = Path(os.environ.get("DOWNLOAD_DIR", "/tmp/labi-downloads"))
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

ALLOWED_EXTENSIONS = {"pdf"}


def erlaubte_datei(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route("/")
def index():
    """Upload-Formular."""
    return render_template("index.html")


@app.route("/health")
def health():
    """Health-Check für Render-Monitoring."""
    return {"status": "ok", "timestamp": datetime.utcnow().isoformat()}


@app.route("/upload", methods=["POST"])
def upload():
    """PDF hochladen und XLSX generieren."""
    if "pdf_file" not in request.files:
        flash("Keine Datei hochgeladen", "error")
        return redirect(url_for("index"))

    file = request.files["pdf_file"]
    if file.filename == "":
        flash("Keine Datei ausgewählt", "error")
        return redirect(url_for("index"))

    if not erlaubte_datei(file.filename):
        flash("Nur PDF-Dateien erlaubt", "error")
        return redirect(url_for("index"))

    # Datei sicher speichern
    original_name = secure_filename(file.filename)
    unique_id = uuid.uuid4().hex[:8]
    pdf_pfad = UPLOAD_DIR / f"{unique_id}_{original_name}"
    file.save(pdf_pfad)
    logger.info(f"PDF gespeichert: {pdf_pfad} ({pdf_pfad.stat().st_size} bytes)")

    # PDF parsen
    try:
        ergebnis = parse_pdf(str(pdf_pfad))
    except Exception as e:
        logger.exception("Parser-Fehler")
        flash(f"Fehler beim Lesen des PDFs: {e}", "error")
        return redirect(url_for("index"))

    if ergebnis.fehler:
        # Schwerer Fehler — kein Download möglich
        logger.warning(f"PDF {original_name}: {ergebnis.fehler}")
        return render_template(
            "fehler.html",
            fehler=ergebnis.fehler,
            warnungen=ergebnis.warnungen,
            header_text=ergebnis.header_text[:500],
        ), 400

    # XLSX generieren
    xlsx_name = original_name.rsplit(".", 1)[0] + "_plancraft.xlsx"
    xlsx_pfad = DOWNLOAD_DIR / f"{unique_id}_{xlsx_name}"
    stats = schreibe_xlsx(ergebnis, str(xlsx_pfad))
    logger.info(f"XLSX generiert: {xlsx_pfad} ({stats['anzahl_positionen']} Positionen)")

    # Erfolgs-Seite mit Download-Link
    return render_template(
        "erfolg.html",
        download_name=f"{unique_id}_{xlsx_name}",
        original_name=original_name,
        stats=stats,
        warnungen=ergebnis.warnungen,
        positionen=ergebnis.positionen,
    )


@app.route("/download/<path:filename>")
def download(filename):
    """Generierte XLSX herunterladen."""
    sicher_name = secure_filename(filename)
    xlsx_pfad = DOWNLOAD_DIR / sicher_name
    if not xlsx_pfad.exists():
        flash("Datei nicht mehr verfügbar", "error")
        return redirect(url_for("index"))

    return send_file(
        str(xlsx_pfad),
        as_attachment=True,
        download_name=sicher_name.split("_", 1)[1] if "_" in sicher_name else sicher_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("DEBUG", "false").lower() == "true"
    logger.info(f"Starte LV2Plancraft auf Port {port} (debug={debug})")
    app.run(host="0.0.0.0", port=port, debug=debug)
