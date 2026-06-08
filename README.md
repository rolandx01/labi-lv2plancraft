# LV → Plancraft Konverter (Webapp, deterministisch, ohne KI)

PDF-Leistungsverzeichnis hochladen → 5-Spalten-Plancraft-XLSX runterladen.
Kein ChatGPT, keine API, keine Token-Kosten. Läuft auf Render.com.

**Live-Demo:** https://labi-lv2plancraft.onrender.com (nach Deployment)

## Was es kann

- Liest GAEB-orientierte LV-PDFs (typische Architekten-/Ausschreiber-Formate)
- Extrahiert Positionen, Kurztexte, Langtexte, Mengen, Einheiten, Preise
- Erzeugt eine XLSX im 5-Spalten-Plancraft-Import-Format
- Web-Upload per Browser — Doppelklick-Ware für Endkunden

## Was es (noch) nicht kann

- Gescannte PDFs (Bild statt Text) → bräuchte OCR (Tesseract)
- Reine GAEB-Dateien (.x83, .d81) → anderes Tool
- Andere LV-Software (ORCA AVA, Sidoun, etc.) → Format noch nicht bekannt

## Lokal starten

```bash
cd /home/ki/hermes-workspace/labi-tool
./start.sh
# → http://localhost:5000
```

## Deployment (Render.com)

Render liest `render.yaml` automatisch und deployt vom `main`-Branch.

```bash
# 1. GitHub-Repo erstellen
gh repo create labi-lv2plancraft --public --source=. --remote=origin --push

# 2. Auf https://render.com → "New +" → "Blueprint" → Repo verbinden
# → Render erkennt render.yaml und deployed automatisch
```

## Dateien

- `app.py` — Flask-App (gunicorn-kompatibel)
- `lv_parser.py` — PDF-Parser (Pattern-Matching, keine KI)
- `xlsx_writer.py` — XLSX-Generator (openpyxl)
- `templates/` — UI (index, erfolg, fehler)
- `requirements.txt` — Python-Dependencies
- `Procfile` — gunicorn-Startbefehl (Render/Fly.io/Heroku-Standard)
- `render.yaml` — Render-Blueprint-Config
- `start.sh` — Lokaler Start
- `test_e2e.sh` — End-to-End-Test

## Tech-Stack

- **Backend:** Flask 3 + gunicorn
- **PDF-Parsing:** pdfplumber (genauer als pypdf für Tabellen)
- **XLSX:** openpyxl
- **Hosting:** Render.com (kostenloser Tier, 750h/Monat)

## Bekannte Einschränkungen

- Render Free-Tier: schläft nach 15 Min Inaktivität ein, Aufwachen dauert ~30 Sek
- Für 24/7-Verfügbarkeit: Render Starter Plan ($7/Monat)
- Labi-Files werden in /tmp gespeichert (auf Render: read-only Filesystem, /tmp ist beschreibbar)
- Cleanup läuft nicht automatisch — bei viel Traffic müsste man /tmp regelmäßig leeren
