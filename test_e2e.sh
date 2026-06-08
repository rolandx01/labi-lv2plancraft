#!/bin/bash
# End-to-End-Test der Webapp mit einem Mock-LV-PDF
set -e
cd "$(dirname "$0")"
source venv/bin/activate

echo "=== Generiere Test-PDF ==="
python3 - <<'PYEOF'
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm

c = canvas.Canvas("/tmp/test_lv.pdf", pagesize=A4)
width, height = A4

# Deckblatt
c.setFont("Helvetica-Bold", 16)
c.drawString(2*cm, height - 3*cm, "Test-LV")
c.setFont("Helvetica", 10)
c.drawString(2*cm, height - 4*cm, "MBS 128-132 LU-Ru")
c.showPage()

# Header
c.setFont("Helvetica-Bold", 10)
c.drawString(2*cm, height - 2*cm, "Pos Bezeichnung EP GP")
c.showPage()

# Positionen
c.setFont("Helvetica", 10)
y = height - 2*cm
for pos_nr, kurz, eigensch, menge, einh in [
    ("10.20.10", "Arbeitsgerüst aufstellen, LK 3",
     ["Höhe über Gelände : bis 10 m", "Breite : mind. 0,60 m", "Lastklasse : 3"],
     "100,00", "m"),
    ("10.20.20", "Schutzdach herstellen",
     ["Breite : 1,50 m"],
     "1,00", "psch"),
    ("10.20.30", "Leiteraufstieg einbauen",
     ["Höhe über Gelände : bis 10 m"],
     "2,00", "Stk"),
]:
    c.setFont("Helvetica-Bold", 10)
    c.drawString(2*cm, y, f"{pos_nr}  {kurz}")
    y -= 0.5*cm
    c.setFont("Helvetica", 9)
    for e in eigensch:
        c.drawString(2.5*cm, y, e)
        y -= 0.4*cm
    c.setFont("Helvetica", 10)
    c.drawString(2*cm, y, f"  {menge} {einh} ............,..... EUR ............,..... EUR")
    y -= 1.0*cm
c.save()
print("OK: /tmp/test_lv.pdf")
PYEOF

echo ""
echo "=== Teste Parser direkt ==="
python3 -c "
from lv_parser import parse_pdf
r = parse_pdf('/tmp/test_lv.pdf')
print(f'Positionen: {len(r.positionen)} | Fehler: {r.fehler}')
for p in r.positionen:
    print(f'  [{p.pos_nr}] {p.kurztext[:40]:40s} | {p.menge:>6} {p.einheit}')
"

echo ""
echo "=== Server starten, Upload testen, Server stoppen ==="
source venv/bin/activate
python3 app.py &
SERVER_PID=$!
sleep 2

echo "Uploading..."
DOWNLOAD_NAME=$(curl -sS -F "pdf_file=@/tmp/test_lv.pdf" http://localhost:5000/upload | grep -oE '/download/[a-f0-9]+_[^"]+\.xlsx' | head -1)
echo "Download-URL: $DOWNLOAD_NAME"

if [ -n "$DOWNLOAD_NAME" ]; then
    curl -sS -o /tmp/test_output.xlsx "http://localhost:5000${DOWNLOAD_NAME}"
    echo "Download: /tmp/test_output.xlsx"
    python3 -c "
from openpyxl import load_workbook
wb = load_workbook('/tmp/test_output.xlsx')
ws = wb.active
print(f'  Excel: {ws.max_row-2} Positionen, {ws.max_column} Spalten')
"
else
    echo "FEHLER: Kein Download-Link erhalten"
fi

kill $SERVER_PID 2>/dev/null
wait $SERVER_PID 2>/dev/null

echo ""
echo "=== Fertig ==="
