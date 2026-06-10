"""
app.py - Flask Webapp für LV → Plancraft Konverter

Akzeptiert sowohl PDF- als auch XLSX-Uploads:
- PDF: wird mit Pattern-Matching geparst (kein OCR, kein KI)
- XLSX: wird je nach erkanntem Format konvertiert
  - Plancraft-Format (5 Spalten): Einheiten normalisieren
  - Förderantrag-Format (9 Spalten): Spalten mappen + normalisieren

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
from xlsx_parser import parse_xlsx
from xlsx_writer import schreibe_xlsx
from validator import validiere_xlsx

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

ALLOWED_EXTENSIONS = {"pdf", "xlsx"}


def erlaubte_datei(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def erlaubte_endung(filename: str) -> str:
    """Gibt die Dateiendung in Kleinbuchstaben zurück (pdf oder xlsx)."""
    return filename.rsplit(".", 1)[1].lower() if "." in filename else ""


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
    """PDF oder XLSX hochladen und Plancraft-XLSX generieren."""
    if "lv_file" not in request.files:
        flash("Keine Datei hochgeladen", "error")
        return redirect(url_for("index"))

    file = request.files["lv_file"]
    if file.filename == "":
        flash("Keine Datei ausgewählt", "error")
        return redirect(url_for("index"))

    if not erlaubte_datei(file.filename):
        flash("Nur PDF- und XLSX-Dateien erlaubt", "error")
        return redirect(url_for("index"))

    endung = erlaubte_endung(file.filename)

    # Datei sicher speichern
    original_name = secure_filename(file.filename)
    unique_id = uuid.uuid4().hex[:8]
    input_pfad = UPLOAD_DIR / f"{unique_id}_{original_name}"
    file.save(input_pfad)
    logger.info(f"Datei gespeichert: {input_pfad} ({input_pfad.stat().st_size} bytes, {endung})")

    # Parsen je nach Format
    if endung == "pdf":
        try:
            ergebnis = parse_pdf(str(input_pfad))
        except Exception as e:
            logger.exception("PDF-Parser-Fehler")
            flash(f"Fehler beim Lesen des PDFs: {e}", "error")
            return redirect(url_for("index"))
        erkanntes_format = "pdf"
    elif endung == "xlsx":
        try:
            ergebnis = parse_xlsx(str(input_pfad))
        except Exception as e:
            logger.exception("XLSX-Parser-Fehler")
            flash(f"Fehler beim Lesen der XLSX: {e}", "error")
            return redirect(url_for("index"))
        erkanntes_format = ergebnis.erkanntes_format
    else:
        flash(f"Unbekanntes Format: {endung}", "error")
        return redirect(url_for("index"))

    if ergebnis.fehler:
        # Schwerer Fehler — kein Download möglich
        logger.warning(f"Datei {original_name}: {ergebnis.fehler}")
        return render_template(
            "fehler.html",
            fehler=ergebnis.fehler,
            warnungen=ergebnis.warnungen,
            header_text=getattr(ergebnis, "header_text", "")[:500],
        ), 400

    # Plancraft-XLSX generieren
    xlsx_name = original_name.rsplit(".", 1)[0] + "_plancraft.xlsx"
    xlsx_pfad = DOWNLOAD_DIR / f"{unique_id}_{xlsx_name}"
    stats = schreibe_xlsx(ergebnis, str(xlsx_pfad))
    logger.info(f"XLSX generiert: {xlsx_pfad} ({stats['anzahl_positionen']} Positionen, Format: {erkanntes_format})")

    # FIX 3: Pre-Validierung der geschriebenen XLSX, bevor User sie runterladen kann.
    # Hintergrund: Plancrafts "irgendein Problem aufgetaucht" ist kryptisch. Wir prüfen
    # jetzt VOR dem Download, ob die XLSX alle Anforderungen erfüllt.
    validierung = validiere_xlsx(str(xlsx_pfad))
    logger.info(
        f"Validierung: status={validierung.status}, "
        f"probleme={len(validierung.probleme)}, "
        f"warnungen={len(validierung.warnungen)}"
    )

    # Erfolgs-Seite mit Download-Link + Validierungs-Status
    return render_template(
        "erfolg.html",
        download_name=f"{unique_id}_{xlsx_name}",
        original_name=original_name,
        stats=stats,
        warnungen=ergebnis.warnungen,
        positionen=ergebnis.positionen,
        erkanntes_format=erkanntes_format,
        validierung=validierung.to_dict(),
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
