"""
Raport Kasowy – generator XLSX i PDF zgodny z importem do Symfonia F&K
======================================================================
Logika biznesowa:
  • Faktury (JPK_FA / KSeF):
      - sprzedaż (firma = sprzedawca)  → KP
      - zakup    (firma = nabywca)      → KW
  • Wyciąg bankowy (PDF) – tylko wybrane operacje:
      wypłata/wpłata bankomat, przelew wewnętrzny/własny/między kontami
      kwota ujemna (debet) → KP (zasilenie kasy)
      kwota dodatnia (credit) → KW (rozchód kasy do banku)
  • Lista płac / diety → KW
"""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
from pathlib import Path
from typing import Literal, Optional

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Side, Border
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

DocType = Literal["KP", "KW"]


# ============================================================================
# Model danych
# ============================================================================

@dataclass(order=True)
class KasaRecord:
    data: date
    typ: DocType
    numer_dokumentu: str
    kontrahent: str
    nip: str
    kwota: Decimal
    lp: int = field(default=0, compare=False)

    def kwota_str(self) -> str:
        q = self.kwota.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        return str(q).replace(".", ",")

    def data_str(self) -> str:
        return self.data.strftime("%Y-%m-%d")


# ============================================================================
# Parsery XML – JPK_FA i KSeF
# ============================================================================

_GOTOWKA_RE = re.compile(r"got[oó]wk[ai]|cash|got\.", re.IGNORECASE | re.UNICODE)

# KSeF FA(3): FormaPlatnosci 1=gotówka, 2=karta, 3=bon, 4=czek, 5=kredyt, 6=przelew
_KSEF_GOTOWKA_CODES = {"1", "gotówka", "gotowka", "cash"}

_FORMA_FIELDS = ["P_19A", "FormaPlatnosci", "SposobPlatnosci"]


def _get_ns(root: ET.Element) -> tuple[str, dict]:
    m = re.match(r"\{([^}]+)\}", root.tag)
    uri = m.group(1) if m else ""
    return uri, {"tns": uri} if uri else {}


def _all_text(elem: ET.Element) -> str:
    return " ".join(
        t for node in elem.iter()
        for t in [node.text or "", node.tail or ""]
        if t.strip()
    )


def _find_first(elem: ET.Element, ns: dict, *xpaths: str) -> str:
    for xpath in xpaths:
        v = elem.findtext(xpath, namespaces=ns) or ""
        if v.strip():
            return v.strip()
    return ""


def _nip_clean(v: str) -> str:
    return re.sub(r"[^0-9]", "", v)


def _kwota_dec(v: str) -> Decimal:
    try:
        return abs(Decimal(v.replace(",", ".")))
    except InvalidOperation:
        return Decimal("0")


# ---------------------------------------------------------------------------
# Wyznaczenie kierunku faktury: KP (sprzedaż) vs KW (zakup)
# ---------------------------------------------------------------------------

def _kierunek_faktury(
    faktura: ET.Element,
    ns: dict,
    firma_nip: str,
    fmt: str,
) -> DocType:
    """
    Porównuje NIP firmy z NIP sprzedawcy i nabywcy na fakturze.
    JPK_FA v4:  P_5B = NIP sprzedawcy,  P_4B = NIP nabywcy
    JPK_FA old: P_5B lub Podmiot1 NIP   vs  P_3A
    KSeF:       Podmiot1/NIP = sprzedawca, Podmiot2/NIP = nabywca
    Jeśli nie można ustalić → domyślnie KP (sprzedaż).
    """
    if not firma_nip:
        return "KP"

    fn = _nip_clean(firma_nip)

    if fmt == "ksef":
        nip_sprzedawcy = _nip_clean(_find_first(
            faktura, ns,
            "tns:Podmiot1/tns:DaneIdentyfikacyjne/tns:NIP",
            ".//tns:Podmiot1//tns:NIP",
        ))
        nip_nabywcy = _nip_clean(_find_first(
            faktura, ns,
            "tns:Podmiot2/tns:DaneIdentyfikacyjne/tns:NIP",
            ".//tns:Podmiot2//tns:NIP",
        ))
    else:
        # JPK_FA v4: P_5B = sprzedawca, P_4B = nabywca
        nip_sprzedawcy = _nip_clean(_find_first(faktura, ns, "tns:P_5B", "P_5B"))
        nip_nabywcy    = _nip_clean(_find_first(faktura, ns, "tns:P_4B", "P_4B", "tns:P_3A", "P_3A"))

    if fn and nip_sprzedawcy and fn == nip_sprzedawcy:
        return "KP"   # firma jest sprzedawcą → wpływ gotówki
    if fn and nip_nabywcy and fn == nip_nabywcy:
        return "KW"   # firma jest nabywcą → rozchód gotówki

    return "KP"  # domyślnie sprzedaż


# ---------------------------------------------------------------------------
# Wyodrębnienie pól faktury wspólnych dla JPK_FA i KSeF
# ---------------------------------------------------------------------------

def _extract_faktura(
    faktura: ET.Element,
    ns: dict,
    fmt: str,
    firma_nip: str,
) -> Optional[KasaRecord]:

    # Data wystawienia
    data_txt = _find_first(
        faktura, ns,
        "tns:P_1", "P_1", "tns:Fa/tns:P_1", ".//tns:P_1",
    )
    try:
        dok_date = datetime.strptime(data_txt[:10], "%Y-%m-%d").date()
    except (ValueError, IndexError):
        return None

    # Numer faktury: JPK_FA v4 → P_2A; KSeF → Fa/P_2; starsze → P_2
    numer = _find_first(
        faktura, ns,
        "tns:P_2A", "P_2A",
        "tns:Fa/tns:P_2", ".//tns:P_2",
        "tns:P_2", "P_2",
        "tns:NrFaktury",
    ) or "?"

    # Kierunek: KP / KW
    kierunek = _kierunek_faktury(faktura, ns, firma_nip, fmt)

    # Kontrahent – zależy od kierunku
    if kierunek == "KP":
        # sprzedaż → kontrahent to NABYWCA
        kontrahent = _find_first(
            faktura, ns,
            "tns:P_3C", "P_3C",                                      # JPK_FA v4
            "tns:Podmiot2/tns:DaneIdentyfikacyjne/tns:Nazwa",
            ".//tns:Podmiot2//tns:Nazwa",                             # KSeF
            "tns:P_3B", "P_3B", "tns:NabywcaNazwa",                  # starsze
        )
        nip = _nip_clean(_find_first(
            faktura, ns,
            "tns:P_4B", "P_4B",
            "tns:Podmiot2/tns:DaneIdentyfikacyjne/tns:NIP",
            ".//tns:Podmiot2//tns:NIP",
            "tns:P_3A", "P_3A", "tns:NabywcaNIP",
        ))
    else:
        # zakup → kontrahent to SPRZEDAWCA
        kontrahent = _find_first(
            faktura, ns,
            "tns:P_3A", "P_3A",                                      # JPK_FA v4 sprzedawca
            "tns:Podmiot1/tns:DaneIdentyfikacyjne/tns:Nazwa",
            ".//tns:Podmiot1//tns:Nazwa",                             # KSeF
            "tns:P_3B", "P_3B",
        )
        nip = _nip_clean(_find_first(
            faktura, ns,
            "tns:P_5B", "P_5B",
            "tns:Podmiot1/tns:DaneIdentyfikacyjne/tns:NIP",
            ".//tns:Podmiot1//tns:NIP",
        ))

    # Kwota P_15
    kwota = _kwota_dec(_find_first(
        faktura, ns,
        "tns:P_15", "P_15", "tns:Fa/tns:P_15", ".//tns:P_15",
        ".//tns:KwotaNaleznosci", ".//tns:WartoscFaktury",
    ) or "0")

    return KasaRecord(
        data=dok_date,
        typ=kierunek,
        numer_dokumentu=numer,
        kontrahent=kontrahent,
        nip=nip,
        kwota=kwota,
    )


# ---------------------------------------------------------------------------
# parse_jpk_fa
# ---------------------------------------------------------------------------

def _is_gotowka_jpk(faktura: ET.Element, ns: dict) -> bool:
    for f in _FORMA_FIELDS:
        v = faktura.findtext(f"tns:{f}", namespaces=ns) or faktura.findtext(f) or ""
        if v and (_GOTOWKA_RE.search(v) or v.strip() in _KSEF_GOTOWKA_CODES):
            return True
    for xpath in (".//tns:FormaPlatnosci", ".//FormaPlatnosci",
                  "tns:Fa/tns:Platnosc/tns:FormaPlatnosci"):
        v = faktura.findtext(xpath, namespaces=ns) or ""
        if v.strip() in _KSEF_GOTOWKA_CODES or _GOTOWKA_RE.search(v):
            return True
    return bool(_GOTOWKA_RE.search(_all_text(faktura)))


def parse_jpk_fa(
    xml_path: str | Path,
    only_cash: bool = True,
    firma_nip: str = "",
) -> tuple[list[KasaRecord], dict]:
    tree = ET.parse(xml_path)
    root = tree.getroot()
    uri, ns = _get_ns(root)

    # NIP firmy z Podmiot1 jeśli nie podany
    if not firma_nip:
        firma_nip = _nip_clean(
            root.findtext(".//tns:Podmiot1//tns:NIP", namespaces=ns)
            or root.findtext(".//Podmiot1//NIP")
            or ""
        )

    faktury = (
        root.findall(".//tns:Faktura", ns)
        or root.findall("tns:Faktura", ns)
        or root.findall("Faktura")
    )

    records, skipped = [], 0
    has_forma = False

    for faktura in faktury:
        # sprawdź czy jest pole formy płatności
        forma_val = any(
            faktura.findtext(f"tns:{f}", namespaces=ns) or faktura.findtext(f)
            for f in _FORMA_FIELDS
        )
        if forma_val:
            has_forma = True

        if only_cash and not _is_gotowka_jpk(faktura, ns):
            skipped += 1
            continue

        rec = _extract_faktura(faktura, ns, "jpk_fa", firma_nip)
        if rec:
            records.append(rec)

    return records, {
        "total": len(records) + skipped,
        "cash": len(records),
        "skipped": skipped,
        "namespace": uri,
        "only_cash": only_cash,
        "has_forma_platnosci": has_forma,
        "firma_nip": firma_nip,
    }


# ---------------------------------------------------------------------------
# parse_ksef
# ---------------------------------------------------------------------------

def parse_ksef(
    xml_path: str | Path,
    only_cash: bool = True,
    firma_nip: str = "",
) -> tuple[list[KasaRecord], dict]:
    tree = ET.parse(xml_path)
    root = tree.getroot()
    uri, ns = _get_ns(root)

    local_root = root.tag.split("}")[-1].lower() if "}" in root.tag else root.tag.lower()
    if local_root == "faktura":
        faktury = [root]
    else:
        faktury = (
            root.findall(".//tns:Faktura", ns)
            or root.findall("tns:Faktura", ns)
            or root.findall("Faktura")
        )

    records, skipped = [], 0

    for faktura in faktury:
        forma = _find_first(
            faktura, ns,
            "tns:Fa/tns:Platnosc/tns:FormaPlatnosci",
            "Fa/Platnosc/FormaPlatnosci",
            ".//tns:FormaPlatnosci",
            ".//FormaPlatnosci",
        )
        is_cash = forma in _KSEF_GOTOWKA_CODES or bool(_GOTOWKA_RE.search(forma))
        if only_cash and not is_cash:
            skipped += 1
            continue

        rec = _extract_faktura(faktura, ns, "ksef", firma_nip)
        if rec:
            records.append(rec)

    return records, {
        "total": len(records) + skipped,
        "cash": len(records),
        "skipped": skipped,
        "namespace": uri,
        "only_cash": only_cash,
        "has_forma_platnosci": True,
        "firma_nip": firma_nip,
    }


# ---------------------------------------------------------------------------
# parse_xml_faktura – auto-detekcja formatu
# ---------------------------------------------------------------------------

def parse_xml_faktura(
    xml_path: str | Path,
    only_cash: bool = True,
    firma_nip: str = "",
) -> tuple[list[KasaRecord], dict, str]:
    """
    Auto-detekcja JPK_FA vs KSeF.
    KSeF: namespace 13775/12648/wzor/2025/2023/2021 lub root=<Faktura>
    """
    tree = ET.parse(xml_path)
    root = tree.getroot()
    tag   = root.tag.lower()
    local = root.tag.split("}")[-1].lower() if "}" in root.tag else root.tag.lower()

    ksef_hints = ["13775", "12648", "wzor/2025", "wzor/2023", "wzor/2021/11", "wzor/2022/01"]
    fmt = "ksef" if (any(h in tag for h in ksef_hints) or local == "faktura") else "jpk_fa"

    if fmt == "ksef":
        recs, diag = parse_ksef(xml_path, only_cash, firma_nip)
    else:
        recs, diag = parse_jpk_fa(xml_path, only_cash, firma_nip)

    return recs, diag, fmt


# ============================================================================
# Parser wyciągu bankowego (PDF)
# ============================================================================

# Operacje kasowe – wpłata/wypłata gotówkowa lub przelew wewnętrzny
_WB_OPERACJE = re.compile(
    r"wypłat[a-z]*\s+(?:w\s+)?bankomat[a-z]*"
    r"|wpłat[a-z]*\s+(?:w\s+|do\s+)?bankomat[a-z]*"
    r"|ATM\s+(?:withdrawal|deposit|cash)"
    r"|transakcja\s+wewnętrzn[a-z]*"
    r"|przelew\s+wewnętrzn[a-z]*"
    r"|przelew\s+własn[a-z]*"
    r"|przelew\s+między\s+kontami"
    r"|operacja\s+między\s+kontami"
    r"|przelew\s+wł\.",
    re.IGNORECASE | re.UNICODE,
)

# Opis przyjazny dla użytkownika (wyodrębniony z linii)
_WB_OPIS = {
    "wypłata.*bankomat":          "Wypłata z bankomatu",
    "wpłata.*bankomat":           "Wpłata do bankomatu",
    "ATM.*withdrawal":            "Wypłata z bankomatu",
    "ATM.*deposit":               "Wpłata do bankomatu",
    "transakcja.*wewnętrzn":      "Transakcja wewnętrzna",
    "przelew.*wewnętrzn":         "Przelew wewnętrzny",
    "przelew.*własn|przelew.*wł": "Przelew własny",
    "przelew.*między.*kontami":   "Przelew między kontami",
    "operacja.*między.*kontami":  "Operacja między kontami",
}

_DATE_RE   = re.compile(r"\b(\d{4}[-/]\d{2}[-/]\d{2}|\d{2}[./]\d{2}[./]\d{4})\b")
# Kwota z opcjonalnym minusem/plusem przed lub po
_AMOUNT_RE = re.compile(r"([+-]?\s*\d[\d\s]*[,.]?\d{0,2})\s*(?:PLN|zł)?", re.IGNORECASE)
_SIGN_RE   = re.compile(r"(?:^|[\s;,])([+-])\s*[\d]")


def _parse_date(txt: str) -> Optional[date]:
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%d.%m.%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(txt.strip(), fmt).date()
        except ValueError:
            continue
    return None


def _parse_signed_amount(line: str) -> tuple[Decimal, Optional[str]]:
    """
    Zwraca (|kwota|, znak) gdzie znak = '+' / '-' / None.
    Kwota ujemna (debet, wypłata z konta) → KP (gotówka wchodzi do kasy).
    Kwota dodatnia (credit, wpłata na konto) → KW (gotówka wychodzi z kasy).
    """
    amt_match = _AMOUNT_RE.search(line)
    if not amt_match:
        return Decimal("0"), None

    raw = amt_match.group(1).replace(" ", "")
    sign = None
    if raw.startswith("-"):
        sign = "-"
        raw = raw[1:]
    elif raw.startswith("+"):
        sign = "+"
        raw = raw[1:]
    else:
        # sprawdź czy przed kwotą jest jawny znak
        s = _SIGN_RE.search(line[:amt_match.start() + 5])
        if s:
            sign = s.group(1)

    try:
        kwota = abs(Decimal(raw.replace(",", ".")))
    except InvalidOperation:
        kwota = Decimal("0")

    return kwota, sign


def _opis_operacji(context: str) -> str:
    """Zwraca przyjazny opis operacji bankowej."""
    for pattern, opis in _WB_OPIS.items():
        if re.search(pattern, context, re.IGNORECASE | re.UNICODE):
            return opis
    return "Operacja kasowa"


def parse_bank_pdf(
    pdf_path: str | Path,
    nr_prefix: str = "BNK",
) -> list[KasaRecord]:
    """
    Przetwarza wyciąg bankowy (PDF).

    Obsługiwane opisy operacji:
      wypłata/wpłata w bankomacie, transakcja wewnętrzna,
      przelew wewnętrzny/własny/między kontami, operacja między kontami

    Logika znaku:
      kwota ujemna (debet)  → KP  (pieniądze opuściły konto → zasiliły kasę)
      kwota dodatnia (credit) → KW  (pieniądze wróciły na konto → wyszły z kasy)
    """
    records: list[KasaRecord] = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            lines = text.splitlines()

            for i, line in enumerate(lines):
                date_match = _DATE_RE.search(line)
                if not date_match:
                    continue
                dok_date = _parse_date(date_match.group())
                if not dok_date:
                    continue

                # kontekst: bieżąca linia + poprzednia + następna
                prev_line = lines[i - 1] if i > 0 else ""
                next_line = lines[i + 1] if i + 1 < len(lines) else ""
                context   = prev_line + " " + line + " " + next_line

                if not _WB_OPERACJE.search(context):
                    continue

                kwota, sign = _parse_signed_amount(line)

                # Gdy brak jawnego znaku – próbuj z kontekstu
                if sign is None:
                    kwota2, sign2 = _parse_signed_amount(context)
                    if sign2:
                        sign = sign2
                        if kwota == Decimal("0"):
                            kwota = kwota2

                # Kierunek: debet (ujemny) → KP, credit (dodatni) → KW
                # Gdy brak znaku – rozstrzygnij po słowie kluczowym
                if sign == "-":
                    typ: DocType = "KP"
                elif sign == "+":
                    typ = "KW"
                elif re.search(r"wypłat[a-z]*\s+(?:w\s+)?bankomat", context, re.I):
                    typ = "KP"
                elif re.search(r"wpłat[a-z]*\s+(?:w\s+|do\s+)?bankomat", context, re.I):
                    typ = "KW"
                else:
                    typ = "KP"  # domyślnie KP gdy brak znaku

                seq   = len(records) + 1
                numer = f"{nr_prefix}/{dok_date.strftime('%Y%m')}/{seq:03d}"
                opis  = _opis_operacji(context)

                records.append(KasaRecord(
                    data=dok_date,
                    typ=typ,
                    numer_dokumentu=numer,
                    kontrahent=opis,
                    nip="",
                    kwota=kwota,
                ))

    return records


# ============================================================================
# Procesor listy płac i diet
# ============================================================================

def process_payroll(
    entries: list[dict],
    pay_date: date,
    nr_prefix: str = "LP",
    collective: bool = False,
) -> list[KasaRecord]:
    """
    Przetwarza listę wypłat gotówkowych.

    Każdy element entries może zawierać:
        - 'nazwisko'     : str
        - 'kwota'        : str | float | Decimal
        - 'typ_wyplaty'  : 'wynagrodzenie' | 'dieta' | ''  (domyślnie 'wynagrodzenie')

    collective=True  → jeden zbiorczy KW
    collective=False → osobny KW per pracownik/dieta
    """
    records: list[KasaRecord] = []

    def _kwota(e: dict) -> Decimal:
        try:
            return abs(Decimal(str(e.get("kwota", 0))).quantize(Decimal("0.01")))
        except InvalidOperation:
            return Decimal("0")

    def _opis(e: dict, idx: int) -> str:
        nazwisko = e.get("nazwisko", f"Pracownik {idx}").strip()
        typ_w    = e.get("typ_wyplaty", "wynagrodzenie").lower()
        if typ_w == "dieta":
            return f"Dieta – {nazwisko}" if nazwisko else f"Dieta #{idx}"
        return nazwisko or f"Pracownik {idx}"

    if collective:
        total = sum(_kwota(e) for e in entries)
        # Jeśli mieszane typy – zbiorczy opis ogólny
        typy = {e.get("typ_wyplaty", "wynagrodzenie").lower() for e in entries}
        if typy == {"dieta"}:
            opis_zb = "Wypłata diet – lista"
            prefix_zb = "DT"
        elif typy == {"wynagrodzenie"} or not typy - {"wynagrodzenie", ""}:
            opis_zb = "Wypłata wynagrodzeń – lista płac"
            prefix_zb = nr_prefix
        else:
            opis_zb = "Wypłata wynagrodzeń i diet – lista"
            prefix_zb = nr_prefix

        records.append(KasaRecord(
            data=pay_date,
            typ="KW",
            numer_dokumentu=f"{prefix_zb}/{pay_date.strftime('%Y%m')}/ZB",
            kontrahent=opis_zb,
            nip="",
            kwota=total,
        ))
    else:
        wynagrodzenia = [(e, i) for i, e in enumerate(entries, 1)
                         if e.get("typ_wyplaty", "wynagrodzenie").lower() != "dieta"]
        diety         = [(e, i) for i, e in enumerate(entries, 1)
                         if e.get("typ_wyplaty", "").lower() == "dieta"]

        # Wynagrodzenia – prefix LP
        for seq, (e, orig_idx) in enumerate(wynagrodzenia, 1):
            records.append(KasaRecord(
                data=pay_date,
                typ="KW",
                numer_dokumentu=f"{nr_prefix}/{pay_date.strftime('%Y%m')}/{seq:03d}",
                kontrahent=_opis(e, orig_idx),
                nip="",
                kwota=_kwota(e),
            ))

        # Diety – prefix DT
        for seq, (e, orig_idx) in enumerate(diety, 1):
            records.append(KasaRecord(
                data=pay_date,
                typ="KW",
                numer_dokumentu=f"DT/{pay_date.strftime('%Y%m')}/{seq:03d}",
                kontrahent=_opis(e, orig_idx),
                nip="",
                kwota=_kwota(e),
            ))

    return records


# ============================================================================
# Główna klasa – Raport Kasowy
# ============================================================================

class RaportKasowy:
    def __init__(self, okres: str = ""):
        self.okres = okres
        self._records: list[KasaRecord] = []

    def dodaj_rekord(self, record: KasaRecord) -> "RaportKasowy":
        self._records.append(record)
        return self

    def _prepare(self) -> list[KasaRecord]:
        sorted_records = sorted(self._records)
        counters: dict[str, int] = {"KP": 0, "KW": 0}
        for i, r in enumerate(sorted_records, 1):
            r.lp = i
            counters[r.typ] += 1
            r.numer_dokumentu = (
                f"{r.typ}/{counters[r.typ]:03d}"
                f"/{r.data.strftime('%m')}/{r.data.strftime('%Y')}"
            )
        return sorted_records

    # -- eksport XLSX -----------------------------------------------------------

    def eksportuj_xlsx(self, sciezka: str | Path) -> Path:
        records = self._prepare()
        wb = Workbook()
        ws = wb.active
        ws.title = "RaportKasowy"

        col_widths = [6, 14, 6, 22, 35, 14, 14]
        for col, w in enumerate(col_widths, 1):
            ws.column_dimensions[get_column_letter(col)].width = w

        font_data    = Font(name="Arial", size=10)
        align_center = Alignment(horizontal="center", vertical="center")
        align_right  = Alignment(horizontal="right",  vertical="center")
        align_left   = Alignment(horizontal="left",   vertical="center")
        fill_kp      = PatternFill("solid", start_color="E8F5E9")
        fill_kw      = PatternFill("solid", start_color="FFF3E0")
        thin         = Side(style="thin", color="CCCCCC")
        border       = Border(left=thin, right=thin, top=thin, bottom=thin)
        aligns       = [align_center, align_center, align_center,
                        align_left, align_left, align_center, align_right]

        for r in records:
            ws.append([r.lp, r.data_str(), r.typ, r.numer_dokumentu,
                        r.kontrahent, r.nip, r.kwota_str()])
            fill    = fill_kp if r.typ == "KP" else fill_kw
            row_num = ws.max_row
            for col_idx, aln in enumerate(aligns, 1):
                cell            = ws.cell(row=row_num, column=col_idx)
                cell.font       = font_data
                cell.alignment  = aln
                cell.fill       = fill
                cell.border     = border
            kwota_cell = ws.cell(row=row_num, column=7)
            try:
                kwota_cell.value          = float(str(r.kwota_str()).replace(",", "."))
                kwota_cell.number_format  = '#,##0.00_-'
            except Exception:
                pass

        out = Path(sciezka)
        wb.save(out)
        return out

    # -- eksport PDF ------------------------------------------------------------

    def eksportuj_pdf(self, sciezka: str | Path) -> Path:
        records = self._prepare()
        out     = Path(sciezka)

        doc = SimpleDocTemplate(
            str(out), pagesize=landscape(A4),
            leftMargin=1.5*cm, rightMargin=1.5*cm,
            topMargin=2*cm, bottomMargin=2*cm,
        )

        styles      = getSampleStyleSheet()
        title_style = ParagraphStyle(
            "title", parent=styles["Title"],
            fontName="Helvetica-Bold", fontSize=14,
            textColor=colors.HexColor("#1a237e"), spaceAfter=4,
        )
        sub_style = ParagraphStyle(
            "sub", parent=styles["Normal"],
            fontName="Helvetica", fontSize=9,
            textColor=colors.gray, spaceAfter=10,
        )

        story   = []
        story.append(Paragraph("RAPORT KASOWY", title_style))
        okres_txt = f"Okres: {self.okres}    " if self.okres else ""
        story.append(Paragraph(
            f"{okres_txt}Wygenerowano: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
            sub_style,
        ))

        headers         = ["Lp.", "Data", "Typ", "Nr dokumentu", "Kontrahent", "NIP", "Kwota (PLN)"]
        col_widths_pdf  = [1.2*cm, 2.5*cm, 1.2*cm, 4.5*cm, 9*cm, 3*cm, 3*cm]

        data_table = [headers]
        total_kp   = Decimal("0")
        total_kw   = Decimal("0")

        for r in records:
            data_table.append([str(r.lp), r.data_str(), r.typ, r.numer_dokumentu,
                                r.kontrahent, r.nip, r.kwota_str()])
            if r.typ == "KP":
                total_kp += r.kwota
            else:
                total_kw += r.kwota

        saldo = total_kp - total_kw
        data_table += [
            ["", "", "", "", "RAZEM", "KP:",    _fmt_dec(total_kp)],
            ["", "", "", "", "",      "KW:",    _fmt_dec(total_kw)],
            ["", "", "", "", "",      "SALDO:", _fmt_dec(saldo)],
        ]

        n  = len(data_table)
        ts = TableStyle([
            ("BACKGROUND",  (0, 0),     (-1, 0),     colors.HexColor("#1a237e")),
            ("TEXTCOLOR",   (0, 0),     (-1, 0),     colors.white),
            ("FONTNAME",    (0, 0),     (-1, 0),     "Helvetica-Bold"),
            ("FONTSIZE",    (0, 0),     (-1, n-4),   8),
            ("VALIGN",      (0, 0),     (-1, -1),    "MIDDLE"),
            ("ALIGN",       (0, 0),     (-1, 0),     "CENTER"),
            ("ALIGN",       (0, 1),     (0, -1),     "CENTER"),
            ("ALIGN",       (1, 1),     (1, -1),     "CENTER"),
            ("ALIGN",       (2, 1),     (2, -1),     "CENTER"),
            ("ALIGN",       (5, 1),     (5, -1),     "CENTER"),
            ("ALIGN",       (6, 1),     (6, -1),     "RIGHT"),
            ("GRID",        (0, 0),     (-1, n-4),   0.3, colors.HexColor("#dddddd")),
            ("LINEABOVE",   (0, 0),     (-1, 0),     1.5, colors.HexColor("#1a237e")),
            ("LINEBELOW",   (0, 0),     (-1, 0),     1.5, colors.HexColor("#1a237e")),
            ("FONTNAME",    (0, n-3),   (-1, -1),    "Helvetica-Bold"),
            ("ALIGN",       (5, n-3),   (6, -1),     "RIGHT"),
            ("LINEABOVE",   (0, n-3),   (-1, n-3),   1, colors.HexColor("#1a237e")),
            ("BACKGROUND",  (0, n-3),   (-1, -1),    colors.HexColor("#e8eaf6")),
            *[("BACKGROUND", (0, i), (-1, i), colors.HexColor("#f5f5f5"))
              for i in range(2, n-3, 2)],
            *[("TEXTCOLOR",  (2, i), (2, i),
               colors.HexColor("#1b5e20") if data_table[i][2] == "KP"
               else colors.HexColor("#bf360c"))
              for i in range(1, n-3)],
        ])

        table = Table(data_table, colWidths=col_widths_pdf, repeatRows=1)
        table.setStyle(ts)
        story.append(table)

        story.append(Spacer(1, 1*cm))
        sign_table = Table(
            [["Sporządził:", "", "Zatwierdził:"],
             ["", "", ""],
             ["_________________________", "", "_________________________"]],
            colWidths=[7*cm, 5*cm, 7*cm],
        )
        sign_table.setStyle(TableStyle([
            ("FONTNAME",   (0, 0), (-1, -1), "Helvetica"),
            ("FONTSIZE",   (0, 0), (-1, -1), 8),
            ("ALIGN",      (0, 0), (-1, -1), "CENTER"),
            ("TOPPADDING", (0, 1), (-1, 1),  14),
        ]))
        story.append(sign_table)
        doc.build(story)
        return out


def _fmt_dec(val: Decimal) -> str:
    return str(val.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)).replace(".", ",")
