"""
xlsx_writer.py - Schreibt ParseResult in Plancraft-Import-XLSX

Plancraft-Import-Format (5 Spalten):
  A: Menge (Zahl)
  B: Einheit (Text: stk, m, m², psch, ...)
  C: Kurztext (Titel)
  D: Langtext (Detail-Beschreibung)
  E: Einheitspreis (optional, € pro Einheit)

Referenz: plancraft_import_dokumente.xlsx
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from typing import Union
from lv_parser import ParseResult, Position
from xlsx_parser import XlsxResult


# Plancraft-Spaltenbreiten (geschätzt, passen für die meisten Angebote)
SPALTEN_BREITEN = {
    "A": 12,   # Menge
    "B": 10,   # Einheit
    "C": 50,   # Kurztext
    "D": 80,   # Langtext
    "E": 14,   # Einheitspreis
}


HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="2C5F8D", end_color="2C5F8D", fill_type="solid")


def schreibe_xlsx(result: Union[ParseResult, XlsxResult], output_pfad: str) -> dict:
    """Schreibt ParseResult in eine Plancraft-kompatible XLSX.

    Args:
        result: ParseResult vom lv_parser
        output_pfad: Zielpfad für .xlsx

    Returns:
        dict mit Statistiken (anzahl_positionen, mit_preis, ohne_preis)
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Positionen"

    # Header schreiben
    # Plancraft-Originalvorlage (plancraft_import_dokumente.xlsx) verwendet EXAKT diese
    # Spaltenüberschriften. Variante wie "Einheitspreis (€)" wurde von Plancraft beim
    # Import offenbar ignoriert (alle Preise = 0,00 €). Daher: exakt diese Schreibweise.
    headers = ["Menge", "Einheit", "Kurztext", "Langtext", "Einheitspreis"]
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Spaltenbreiten setzen
    for col_letter, breite in SPALTEN_BREITEN.items():
        ws.column_dimensions[col_letter].width = breite

    # Header-Zeile einfrieren
    ws.freeze_panes = "A2"

    # Positionen schreiben
    row = 2
    mit_preis = 0
    ohne_preis = 0

    for pos in result.positionen:
        # Menge (leer lassen wenn None, statt 0)
        ws.cell(row=row, column=1, value=pos.menge if pos.menge is not None else "")

        # Einheit: Wenn leer, default "Stk." (Plancrafts Standard-Einheit)
        # Plancraft mag leere Einheiten nicht — setzt sonst auf "1,00 Stk." was oft
        # falsch ist für "m²", "l" etc. Wir geben lieber direkt Stk. mit und Labi
        # ändert es bei Bedarf in Plancraft per Klick.
        if pos.einheit:
            ws.cell(row=row, column=2, value=pos.einheit)
        else:
            ws.cell(row=row, column=2, value="Stk.")

        # Kurztext
        ws.cell(row=row, column=3, value=pos.kurztext)

        # Langtext
        ws.cell(row=row, column=4, value=pos.langtext)

        # Einheitspreis (optional)
        if pos.einheitspreis is not None:
            cell = ws.cell(row=row, column=5, value=pos.einheitspreis)
            cell.number_format = "#,##0.00 €"
            mit_preis += 1
        else:
            ws.cell(row=row, column=5, value="")
            ohne_preis += 1

        # Zeilenhöhe auto-anpassen
        ws.row_dimensions[row].height = None  # Excel macht das automatisch

        # Wrap-Text für Langtext aktivieren
        ws.cell(row=row, column=4).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=row, column=3).alignment = Alignment(wrap_text=True, vertical="top")

        row += 1

    # Zusatz-Info als Kommentar in einer freien Zelle (optional)
    if result.warnungen or result.fehler:
        info_row = row + 2
        ws.cell(row=info_row, column=1, value="Hinweise:").font = Font(bold=True, italic=True)
        for idx, warn in enumerate(result.warnungen, start=1):
            ws.cell(row=info_row + idx, column=1, value=f"⚠ {warn}").font = Font(italic=True, color="B45F06")
        for idx, err in enumerate(result.fehler, start=1):
            ws.cell(row=info_row + idx + len(result.warnungen), column=1, value=f"✗ {err}").font = Font(italic=True, color="CC0000")

    wb.save(output_pfad)

    return {
        "anzahl_positionen": len(result.positionen),
        "mit_preis": mit_preis,
        "ohne_preis": ohne_preis,
        "zeilen_geschrieben": row - 2,
    }


if __name__ == "__main__":
    import sys
    from lv_parser import parse_pdf

    if len(sys.argv) < 3:
        print("Usage: python xlsx_writer.py <pdf-pfad> <xlsx-pfad>")
        sys.exit(1)

    ergebnis = parse_pdf(sys.argv[1])
    stats = schreibe_xlsx(ergebnis, sys.argv[2])
    print(f"XLSX geschrieben: {sys.argv[2]}")
    print(f"  Positionen: {stats['anzahl_positionen']}")
    print(f"  Mit Preis:  {stats['mit_preis']}")
    print(f"  Ohne Preis: {stats['ohne_preis']}")
