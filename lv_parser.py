"""
lv_parser.py - PDF LV-Extraktor (deterministisch, ohne KI)

Erkennt GAEB-orientierte Leistungsverzeichnis-Struktur:
  - Position: XX.XX.XX  Kurztext
  - Langtext: Fließtext + Eigenschafts-Felder ("Höhe über Gelände : 10m")
  - Menge + Einheit: "  10,00 m ............,..... EUR ............,..... EUR"

Wichtig: KEINE KI, keine API-Calls. Reines Pattern-Matching mit Regex.
Wenn ein PDF von dieser Struktur abweicht, wird es als Fehler gemeldet
(nicht stillschweigend Müll produziert).
"""

import re
from dataclasses import dataclass, field
from typing import List, Optional
import pdfplumber


@dataclass
class Position:
    """Eine einzelne LV-Position im Plancraft-Format."""
    menge: Optional[float] = None
    einheit: Optional[str] = None
    kurztext: str = ""
    langtext: str = ""
    einheitspreis: Optional[float] = None
    gesamtpreis: Optional[float] = None
    pos_nr: str = ""
    seite: int = 0  # Zur Nachverfolgung


@dataclass
class ParseResult:
    """Ergebnis eines PDF-Pars."""
    positionen: List[Position] = field(default_factory=list)
    fehler: List[str] = field(default_factory=list)
    warnungen: List[str] = field(default_factory=list)
    header_text: str = ""  # Deckblatt/Anschreiben für Kontext
    seiten_gesamt: int = 0
    positionen_bereich: tuple = (0, 0)  # (start_seite, end_seite)


# --- Regex-Muster ---

# Position: "10.20.80  Konsolen..." oder "10.20.80 Konsolen..."
# (Ziffern-Punkte-Punkte, dann 1+ Spaces, dann Text)
# Wir akzeptieren 1+ Leerzeichen, weil Architekten-PDFs variieren.
# Zusätzlich: nach der Pos-Nr muss ein BEGRIFF folgen (Buchstabe), nicht eine Zahl,
# um False-Positives auf z.B. "10.20 1,5 m" (Menge mit Punkt) zu vermeiden.
RE_POS_NR = re.compile(r"^(\d{1,3}\.\d{1,3}(?:\.\d{1,3})?)\s+([A-Za-zÄÖÜäöüß].+)$")

# Menge + Einheit + EP + GP: "  10,00 m ............,..... EUR ............,..... EUR"
# Erfasst: (1) Menge, (2) Einheit, (3) Einheitspreis (optional), (4) Gesamtpreis (optional)
RE_MENGE_EINHEIT = re.compile(
    r"^\s*"
    r"(?P<menge>\d{1,3}(?:\.\d{3})*(?:,\d+)?|\d+)\s+"  # Menge: 10,00 / 1.000,50 / 100
    r"(?P<einheit>[A-Za-z²³/]{1,8})\s+"                   # Einheit: m / m² / Stk / psch
    r"(?:\.{3,},?\s*EUR\s+)?"
    r"(?:\.{3,},?\s*EUR\s*)?"
    r"(?P<preis>\d{1,3}(?:\.\d{3})*(?:,\d+)?)\s*"
    r"EUR?"
    r"\s*$"
)

# Einheiten, die wir akzeptieren
GUELTIGE_EINHEITEN = {
    "stk", "st", "stck", "stück", "stk.",
    "m", "lfm", "m²", "m2", "qm", "m3", "m³",
    "kg", "t",
    "l", "ltr", "ml",
    "h", "std", "min",
    "psch", "psch.", "pausch", "pauschal", "pausch.", "pauschwo", "pauchwo",
    "wo", "tge", "tage", "at", "at.",
    "mwo", "mwo.", "stwo", "stwo.", "stkwo",
    "km", "kwh", "%",
    "lfd", "lfdm", "lfm.",
}


# Eigenschafts-Felder wie "Höhe über Gelände : bis 10 m"
RE_EIGENSCHAFT = re.compile(r"^([A-ZÄÖÜ][\w\s\(\)\-]+?)\s*:\s*(.+)$")


def ist_eigenschafts_feld(zeile: str) -> bool:
    """Prüft, ob eine Zeile wie 'Höhe über Gelände : 10 m' aussieht."""
    if not RE_EIGENSCHAFT.match(zeile.strip()):
        return False
    # Blacklist: Wörter, die KEINE Eigenschaften sind, sondern normaler Text
    blacklist_start = (
        "Wir ", "Sie ", "Der ", "Die ", "Das ", "Alle ", "Bei ",
        "Pos ", "Position ", "LV ", "Seite ", "Menge ", "Betrag ",
    )
    for w in blacklist_start:
        if zeile.strip().startswith(w):
            return False
    return True


def ist_header_position_header(zeile: str) -> bool:
    """Erkennt die 'Pos Bezeichnung EP'-Tabellenkopf-Zeile."""
    z = zeile.strip().lower()
    return "pos" in z and "bezeichnung" in z and ("ep" in z or "einzelpreis" in z or "gesamt" in z)


def ist_footer_oder_seitenzahl(zeile: str) -> bool:
    """Erkennt Seitenzahlen, LV-Footer etc."""
    z = zeile.strip()
    if re.match(r"^[-–—]?\s*\d{1,3}\s*[-–—]?\s*(von\s+\d+)?$", z):
        return True
    if z in {"", "EUR", "€", "Betrag EUR", "Summe EUR"}:
        return True
    return False


def parse_menge_einheit(zeile: str) -> Optional[tuple]:
    """Versucht Menge + Einheit aus einer Zeile zu extrahieren.
    Gibt (menge, einheit, einheitspreis, gesamtpreis) zurück oder None.
    """
    # Vereinfachte Version: Menge am Zeilenanfang, dann Einheit
    match = re.match(
        r"^\s*"
        r"(\d{1,3}(?:\.\d{3})*(?:,\d+)?|\d+)\s+"  # Menge
        r"([A-Za-z²³/]{1,8})"                       # Einheit
        r"\s*(.*)$",                                 # Rest (für Preise)
        zeile,
    )
    if not match:
        return None

    menge_str, einheit, rest = match.groups()
    einheit_norm = einheit.lower().rstrip(".")

    # Nur gültige Einheiten akzeptieren
    if einheit_norm not in GUELTIGE_EINHEITEN:
        return None

    # Menge zu float
    try:
        menge = float(menge_str.replace(".", "").replace(",", "."))
    except ValueError:
        return None

    # Preise extrahieren (falls vorhanden)
    einheitspreis = None
    gesamtpreis = None
    if rest:
        preise = re.findall(r"(\d{1,3}(?:\.\d{3})*(?:,\d+)?)\s*EUR?", rest)
        if len(preise) >= 1:
            try:
                einheitspreis = float(preise[0].replace(".", "").replace(",", "."))
            except ValueError:
                pass
        if len(preise) >= 2:
            try:
                gesamtpreis = float(preise[1].replace(".", "").replace(",", "."))
            except ValueError:
                pass

    return menge, einheit_norm, einheitspreis, gesamtpreis


def parse_pdf(pdf_pfad: str) -> ParseResult:
    """Parst ein LV-PDF und gibt alle erkannten Positionen zurück.

    Args:
        pdf_pfad: Absoluter Pfad zum PDF

    Returns:
        ParseResult mit positionen, fehler, warnungen
    """
    result = ParseResult()

    try:
        with pdfplumber.open(pdf_pfad) as pdf:
            result.seiten_gesamt = len(pdf.pages)
            alle_zeilen = []  # (seite_nr, zeile)

            for seite_nr, seite in enumerate(pdf.pages, start=1):
                text = seite.extract_text() or ""
                for zeile in text.split("\n"):
                    alle_zeilen.append((seite_nr, zeile))

    except Exception as e:
        result.fehler.append(f"PDF konnte nicht gelesen werden: {e}")
        return result

    # Phase 1: Finde Start des LV-Positionsteils
    # Strategie: Suche nach der Zeile "Pos Bezeichnung EP" oder ähnlich
    pos_start_idx = None
    for idx, (seite, zeile) in enumerate(alle_zeilen):
        if ist_header_position_header(zeile):
            pos_start_idx = idx + 1  # Header-Zeile selbst überspringen
            result.positionen_bereich = (seite, result.seiten_gesamt)
            break

    if pos_start_idx is None:
        # Fallback: Suche nach erster Position mit XX.XX.XX-Muster
        for idx, (seite, zeile) in enumerate(alle_zeilen):
            if RE_POS_NR.match(zeile.strip()):
                pos_start_idx = idx
                result.warnungen.append(
                    "Kein 'Pos Bezeichnung EP'-Header gefunden — "
                    "parse ab erster erkannter Position. "
                    "Bitte prüfe, ob alle Positionen erfasst wurden."
                )
                break

    if pos_start_idx is None:
        result.fehler.append(
            "Keine LV-Positionen erkannt. "
            "Mögliche Ursachen: PDF ist gescannt (OCR nötig), "
            "anderes LV-Format, oder Struktur weicht von GAEB ab."
        )
        # Header trotzdem sammeln
        result.header_text = "\n".join(z for _, z in alle_zeilen[:50])
        return result

    # Phase 2: Parse Positionen
    current_pos: Optional[Position] = None
    current_langtext_zeilen: List[str] = []
    position_erwartet_menge = False  # Nächste Zeile nach Pos-Nr könnte Menge+Einheit sein

    for idx in range(pos_start_idx, len(alle_zeilen)):
        seite, zeile_raw = alle_zeilen[idx]
        zeile = zeile_raw.strip()

        # Leerzeile = Ende der aktuellen Position (Langtext abschließen)
        if not zeile:
            if current_pos and current_langtext_zeilen:
                current_pos.langtext = "\n".join(current_langtext_zeilen).strip()
            current_pos = None
            current_langtext_zeilen = []
            position_erwartet_menge = False
            continue

        # Footer/Seitenzahl überspringen
        if ist_footer_oder_seitenzahl(zeile):
            continue

        # Neue Position erkannt?
        pos_match = RE_POS_NR.match(zeile)
        if pos_match:
            # Vorherige Position abschließen
            if current_pos:
                if current_langtext_zeilen:
                    current_pos.langtext = "\n".join(current_langtext_zeilen).strip()
                result.positionen.append(current_pos)

            pos_nr, kurztext = pos_match.groups()
            current_pos = Position(
                pos_nr=pos_nr,
                kurztext=kurztext.strip(),
                seite=seite,
            )
            current_langtext_zeilen = []
            position_erwartet_menge = True
            continue

        # Wenn wir in einer Position sind
        if current_pos:
            # Versuche IMMER Menge + Einheit zu extrahieren, solange die Zeile
            # nicht bereits als Eigenschaft/Fließtext "verbraucht" ist.
            # Realität: Die Menge-Zeile (XX,XX Einheit ..... EUR ..... EUR)
            # ist immer sehr eindeutig — die erwischen wir auch nach Fließtext.
            menge_einheit = parse_menge_einheit(zeile)
            if menge_einheit:
                current_pos.menge, current_pos.einheit, current_pos.einheitspreis, current_pos.gesamtpreis = menge_einheit
                position_erwartet_menge = False
                continue

            # Eigenschafts-Feld (z.B. "Höhe über Gelände : 10 m")? → in Langtext
            if ist_eigenschafts_feld(zeile):
                current_langtext_zeilen.append(zeile)
                continue

            # Normaler Fließtext-Langtext
            current_langtext_zeilen.append(zeile)

    # Letzte Position abschließen
    if current_pos:
        if current_langtext_zeilen:
            current_pos.langtext = "\n".join(current_langtext_zeilen).strip()
        result.positionen.append(current_pos)

    # Header extrahieren (alles vor pos_start_idx)
    header_zeilen = [z for _, z in alle_zeilen[:pos_start_idx]]
    result.header_text = "\n".join(header_zeilen).strip()

    # Validierung
    if not result.positionen:
        result.fehler.append("Parser hat keine Positionen gefunden.")
    else:
        # Warnung, wenn viele Positionen keine Menge haben
        ohne_menge = sum(1 for p in result.positionen if p.menge is None)
        if ohne_menge > 0:
            anteil = ohne_menge / len(result.positionen) * 100
            if anteil > 20:
                result.warnungen.append(
                    f"{ohne_menge} von {len(result.positionen)} Positionen "
                    f"({anteil:.0f}%) haben keine erkannte Menge. "
                    "Mögliche Ursache: Mengen-Einheit-Zeile hat ungewöhnliches Format."
                )

    return result


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Usage: python lv_parser.py <pdf-pfad>")
        sys.exit(1)

    ergebnis = parse_pdf(sys.argv[1])
    print(f"Seiten: {ergebnis.seiten_gesamt}")
    print(f"Positionen: {len(ergebnis.positionen)}")
    print(f"Fehler: {len(ergebnis.fehler)}")
    print(f"Warnungen: {len(ergebnis.warnungen)}")
    print()
    for f in ergebnis.fehler:
        print(f"FEHLER: {f}")
    for w in ergebnis.warnungen:
        print(f"WARNUNG: {w}")
    print()
    print("Erste 3 Positionen:")
    for p in ergebnis.positionen[:3]:
        print(f"  [{p.pos_nr}] {p.kurztext}")
        print(f"    Menge: {p.menge} {p.einheit} | EP: {p.einheitspreis} | GP: {p.gesamtpreis}")
        print(f"    Langtext: {p.langtext[:80]}...")
        print()
