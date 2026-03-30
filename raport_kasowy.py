"""
Raport Kasowy – generator XLSX i PDF zgodny z importem do Symfonia F&K
======================================================================
Źródła danych:
  • JPK_FA (XML)  – faktury opłacone gotówką → KP
  • Wyciąg bankowy (PDF) – wypłaty z bankomatu → KP, wpłaty do bankomatu → KW
  • Lista płac (dict/CSV)  – wypłaty gotówkowe pracownikom → KW
"""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
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
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
)

# ---------------------------------------------------------------------------
# Typ dokumentu
# ---------------------------------------------------------------------------
DocType = Literal["KP", "KW"]


@dataclass(order=True)
class KasaRecord:
    """Jeden wiersz raportu kasowego."""
    data: date
    typ: DocType
    numer_dokumentu: str
    kontrahent: str
    nip: str
    kwota: Decimal
    lp: int = field(default=0, compare=False)   # wypełniane po sortowaniu

    def kwota_str(self) -> str:
        """Kwota z przecinkiem dziesiętnym (wymóg Symfonia)."""
        q = self.kwota.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        return str(q).replace(".", ",")

    def data_str(self) -> str:
        return self.data.strftime("%Y-%m-%d")


# ---------------------------------------------------------------------------
# Parser JPK_FA
# ---------------------------------------------------------------------------
_JPK_NS = {
    "tns": "http://crd.gov.pl/wzor/2022/11/29/11089/",
    "etd": "http://crd.gov.pl/xml/schematy/dziedzinowe/mf/2018/08/24/eD/DefinicjeTypow/",
}

# Obsługujemy też starszy namespace JPK_FA(3)
_JPK_NS_V3 = {
    "tns": "http://jpk.mf.gov.pl/wzor/2022/02/17/02171/",
    "etd": "http://crd.gov.pl/xml/schematy/dziedzinowe/mf/2018/08/24/eD/DefinicjeTypow/",
}


def _detect_ns(root: ET.Element) -> dict:
    tag = root.tag
    if "11089" in tag or "2022/11/29" in tag:
        return _JPK_NS
    return _JPK_NS_V3


def parse_jpk_fa(xml_path: str | Path) -> list[KasaRecord]:
    """
    Wyodrębnia z JPK_FA faktury opłacone gotówką (FP lub FA z formą
    płatności == 'gotówka'/'gotowka'/'cash').
    Zwraca listę KasaRecord z typem KP.
    """
    tree = ET.parse(xml_path)
    root = tree.getroot()
    ns = _detect_ns(root)

    records: list[KasaRecord] = []

    # Każda faktura to element <tns:Faktura>
    for faktura in root.findall(".//tns:Faktura", ns):
        # Forma płatności
        forma = (faktura.findtext("tns:P_22", default="", namespaces=ns) or "").lower()
        forma += (faktura.findtext("tns:FormaPlatnosci", default="", namespaces=ns) or "").lower()
        gotowka_keywords = {"gotówka", "gotowka", "cash", "got."}
        if not any(k in forma for k in gotowka_keywords):
            continue

        # Data wystawienia
        data_txt = (
            faktura.findtext("tns:P_1", default="", namespaces=ns)
            or faktura.findtext("tns:DataWystawienia", default="", namespaces=ns)
        ).strip()
        try:
            dok_date = datetime.strptime(data_txt[:10], "%Y-%m-%d").date()
        except ValueError:
            continue

        numer = (
            faktura.findtext("tns:P_2", default="", namespaces=ns)
            or faktura.findtext("tns:NrFaktury", default="", namespaces=ns)
            or "?"
        ).strip()

        nabywca = (
            faktura.findtext("tns:P_3B", default="", namespaces=ns)
            or faktura.findtext("tns:NabywcaNazwa", default="", namespaces=ns)
            or ""
        ).strip()

        nip = (
            faktura.findtext("tns:P_3A", default="", namespaces=ns)
            or faktura.findtext("tns:NabywcaNIP", default="", namespaces=ns)
            or ""
        ).strip()
        nip = re.sub(r"[^0-9]", "", nip)   # tylko cyfry

        kwota_txt = (
            faktura.findtext("tns:P_15", default="", namespaces=ns)
            or faktura.findtext("tns:KwotaNaleznosci", default="", namespaces=ns)
            or "0"
        ).replace(",", ".")

        try:
            kwota = abs(Decimal(kwota_txt))
        except Exception:
            kwota = Decimal("0")

        records.append(KasaRecord(
            data=dok_date,
            typ="KP",
            numer_dokumentu=numer,
            kontrahent=nabywca,
            nip=nip,
            kwota=kwota,
        ))

    return records


# ---------------------------------------------------------------------------
# Parser wyciągu bankowego (PDF)
# ---------------------------------------------------------------------------
_ATM_WY = re.compile(
    r"(wypłat[a-z]*\s+z\s+bankomatu|ATM\s+withdrawal|wypłata\s+bankomat)",
    re.IGNORECASE
)
_ATM_WP = re.compile(
    r"(wpłat[a-z]*\s+(do\s+)?bankomatu|ATM\s+deposit|wplata\s+bankomat)",
    re.IGNORECASE
)
_DATE_RE = re.compile(r"\b(\d{4}[-/]\d{2}[-/]\d{2}|\d{2}[./]\d{2}[./]\d{4})\b")
_AMOUNT_RE = re.compile(r"(\d[\d\s]*[,.]?\d*)\s*(PLN|zł)?", re.IGNORECASE)


def _parse_date(txt: str) -> Optional[date]:
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%d.%m.%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(txt.strip(), fmt).date()
        except ValueError:
            continue
    return None


def _parse_amount(txt: str) -> Decimal:
    clean = re.sub(r"[^\d,.]", "", txt)
    clean = clean.replace(",", ".")
    try:
        return abs(Decimal(clean))
    except Exception:
        return Decimal("0")


def parse_bank_pdf(
    pdf_path: str | Path,
    nr_prefix: str = "BNK",
) -> list[KasaRecord]:
    """
    Przetwarza wyciąg bankowy (PDF).
    Wypłaty z bankomatu  → KP  (zasilenie kasy)
    Wpłaty do bankomatu  → KW  (odprowadzenie gotówki do banku)
    """
    records: list[KasaRecord] = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            lines = text.splitlines()

            for i, line in enumerate(lines):
                # szukaj daty w wierszu
                date_match = _DATE_RE.search(line)
                if not date_match:
                    continue
                dok_date = _parse_date(date_match.group())
                if not dok_date:
                    continue

                # szukaj kwoty
                amt_match = _AMOUNT_RE.search(line[date_match.end():])
                kwota = _parse_amount(amt_match.group(1)) if amt_match else Decimal("0")

                # kontekst – bieżący + następny wiersz
                context = line + " " + (lines[i + 1] if i + 1 < len(lines) else "")
                seq = len(records) + 1
                numer = f"{nr_prefix}/{dok_date.strftime('%Y%m')}/{seq:03d}"

                if _ATM_WY.search(context):
                    records.append(KasaRecord(
                        data=dok_date,
                        typ="KP",
                        numer_dokumentu=numer,
                        kontrahent="Wypłata z bankomatu",
                        nip="",
                        kwota=kwota,
                    ))
                elif _ATM_WP.search(context):
                    records.append(KasaRecord(
                        data=dok_date,
                        typ="KW",
                        numer_dokumentu=numer,
                        kontrahent="Wpłata do bankomatu",
                        nip="",
                        kwota=kwota,
                    ))

    return records


# ---------------------------------------------------------------------------
# Procesor listy płac
# ---------------------------------------------------------------------------
def process_payroll(
    entries: list[dict],
    pay_date: date,
    nr_prefix: str = "LP",
    collective: bool = False,
) -> list[KasaRecord]:
    """
    Przetwarza listę płac.

    entries: lista słowników z kluczami:
        - 'nazwisko'  : str
        - 'kwota'     : str | float | Decimal  (brutto lub netto – per ustawienia)
        - 'pesel'     : str (opcjonalnie)

    collective=True  → jeden zbiorczy dokument KW
    collective=False → osobny KW per pracownik
    """
    records: list[KasaRecord] = []

    if collective:
        total = sum(abs(Decimal(str(e.get("kwota", 0))).quantize(Decimal("0.01")))
                    for e in entries)
        records.append(KasaRecord(
            data=pay_date,
            typ="KW",
            numer_dokumentu=f"{nr_prefix}/{pay_date.strftime('%Y%m')}/ZB",
            kontrahent="Wypłata zbiorowa – lista płac",
            nip="",
            kwota=total,
        ))
    else:
        for idx, e in enumerate(entries, 1):
            kwota = abs(Decimal(str(e.get("kwota", 0))).quantize(Decimal("0.01")))
            records.append(KasaRecord(
                data=pay_date,
                typ="KW",
                numer_dokumentu=f"{nr_prefix}/{pay_date.strftime('%Y%m')}/{idx:03d}",
                kontrahent=e.get("nazwisko", f"Pracownik {idx}"),
                nip="",
                kwota=kwota,
            ))

    return records


# ---------------------------------------------------------------------------
# Główna klasa – Raport Kasowy
# ---------------------------------------------------------------------------
class RaportKasowy:
    """
    Agreguje rekordy z różnych źródeł, sortuje chronologicznie
    i eksportuje do XLSX / PDF.
    """

    def __init__(self, okres: str = ""):
        self.okres = okres          # np. "2025-01"
        self._records: list[KasaRecord] = []

    # -- dodawanie danych -------------------------------------------------------

    def dodaj_jpk_fa(self, xml_path: str | Path) -> "RaportKasowy":
        self._records.extend(parse_jpk_fa(xml_path))
        return self

    def dodaj_wyciag_bankowy(
        self, pdf_path: str | Path, nr_prefix: str = "BNK"
    ) -> "RaportKasowy":
        self._records.extend(parse_bank_pdf(pdf_path, nr_prefix))
        return self

    def dodaj_liste_plac(
        self,
        entries: list[dict],
        pay_date: date,
        nr_prefix: str = "LP",
        collective: bool = False,
    ) -> "RaportKasowy":
        self._records.extend(process_payroll(entries, pay_date, nr_prefix, collective))
        return self

    def dodaj_rekord(self, record: KasaRecord) -> "RaportKasowy":
        """Ręczne dodanie pojedynczego rekordu."""
        self._records.append(record)
        return self

    # -- przetwarzanie ----------------------------------------------------------

    def _prepare(self) -> list[KasaRecord]:
        sorted_records = sorted(self._records)
        # Osobne liczniki dla KP i KW
        counters: dict[str, int] = {"KP": 0, "KW": 0}
        for i, r in enumerate(sorted_records, 1):
            r.lp = i
            counters[r.typ] += 1
            nr = counters[r.typ]
            mm = r.data.strftime("%m")
            yyyy = r.data.strftime("%Y")
            r.numer_dokumentu = f"{r.typ}/{nr:03d}/{mm}/{yyyy}"
        return sorted_records

    # -- eksport XLSX -----------------------------------------------------------

    def eksportuj_xlsx(self, sciezka: str | Path) -> Path:
        """
        Generuje plik XLSX gotowy do importu do Symfonia F&K.
        Brak wiersza nagłówkowego. Kolumny: Lp | Data | Typ | Nr | Kontrahent | NIP | Kwota
        """
        records = self._prepare()
        wb = Workbook()
        ws = wb.active
        ws.title = "RaportKasowy"

        # Szerokości kolumn
        col_widths = [6, 14, 6, 22, 35, 14, 14]
        for col, w in enumerate(col_widths, 1):
            ws.column_dimensions[get_column_letter(col)].width = w

        # Style
        font_data = Font(name="Arial", size=10)
        align_center = Alignment(horizontal="center", vertical="center")
        align_right = Alignment(horizontal="right", vertical="center")
        align_left = Alignment(horizontal="left", vertical="center")

        fill_kp = PatternFill("solid", start_color="E8F5E9")  # zielonkawy
        fill_kw = PatternFill("solid", start_color="FFF3E0")  # pomarańczowy

        thin = Side(style="thin", color="CCCCCC")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for r in records:
            row = [
                r.lp,
                r.data_str(),
                r.typ,
                r.numer_dokumentu,
                r.kontrahent,
                r.nip,
                r.kwota_str(),
            ]
            ws.append(row)
            fill = fill_kp if r.typ == "KP" else fill_kw
            row_num = ws.max_row

            aligns = [
                align_center,  # Lp
                align_center,  # Data
                align_center,  # Typ
                align_left,    # Nr
                align_left,    # Kontrahent
                align_center,  # NIP
                align_right,   # Kwota
            ]
            for col_idx, (val, aln) in enumerate(zip(row, aligns), 1):
                cell = ws.cell(row=row_num, column=col_idx)
                cell.font = font_data
                cell.alignment = aln
                cell.fill = fill
                cell.border = border

            # Kwota jako liczba z formatem
            kwota_cell = ws.cell(row=row_num, column=7)
            try:
                kwota_cell.value = float(str(r.kwota_str()).replace(",", "."))
                kwota_cell.number_format = '#,##0.00_-'
            except Exception:
                pass

        out = Path(sciezka)
        wb.save(out)
        return out

    # -- eksport PDF ------------------------------------------------------------

    def eksportuj_pdf(self, sciezka: str | Path) -> Path:
        """
        Generuje raport kasowy w formacie PDF (orientacja pozioma A4).
        """
        records = self._prepare()
        out = Path(sciezka)

        doc = SimpleDocTemplate(
            str(out),
            pagesize=landscape(A4),
            leftMargin=1.5 * cm,
            rightMargin=1.5 * cm,
            topMargin=2 * cm,
            bottomMargin=2 * cm,
        )

        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            "title",
            parent=styles["Title"],
            fontName="Helvetica-Bold",
            fontSize=14,
            textColor=colors.HexColor("#1a237e"),
            spaceAfter=4,
        )
        sub_style = ParagraphStyle(
            "sub",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=9,
            textColor=colors.gray,
            spaceAfter=10,
        )
        cell_style = ParagraphStyle(
            "cell",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=8,
            leading=10,
        )

        story = []

        # Tytuł
        story.append(Paragraph("RAPORT KASOWY", title_style))
        okres_txt = f"Okres: {self.okres}" if self.okres else ""
        gen_txt = f"Wygenerowano: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        story.append(Paragraph(f"{okres_txt}    {gen_txt}", sub_style))

        # Nagłówki tabeli
        headers = ["Lp.", "Data", "Typ", "Nr dokumentu", "Kontrahent", "NIP", "Kwota (PLN)"]
        col_widths_pdf = [1.2 * cm, 2.5 * cm, 1.2 * cm, 4.5 * cm, 9 * cm, 3 * cm, 3 * cm]

        data_table = [headers]
        total_kp = Decimal("0")
        total_kw = Decimal("0")

        for r in records:
            data_table.append([
                str(r.lp),
                r.data_str(),
                r.typ,
                r.numer_dokumentu,
                r.kontrahent,
                r.nip,
                r.kwota_str(),
            ])
            if r.typ == "KP":
                total_kp += r.kwota
            else:
                total_kw += r.kwota

        # Wiersz podsumowania
        saldo = total_kp - total_kw
        data_table.append([
            "", "", "", "", "RAZEM",
            "KP:",
            _fmt_dec(total_kp),
        ])
        data_table.append([
            "", "", "", "", "",
            "KW:",
            _fmt_dec(total_kw),
        ])
        data_table.append([
            "", "", "", "", "",
            "SALDO:",
            _fmt_dec(saldo),
        ])

        # Style tabeli
        n = len(data_table)
        ts = TableStyle([
            # nagłówek
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1a237e")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 8),
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            # dane
            ("FONTNAME", (0, 1), (-1, n - 4), "Helvetica"),
            ("FONTSIZE", (0, 1), (-1, n - 4), 8),
            # kolumny liczbowe – wyrównanie do prawej
            ("ALIGN", (0, 1), (0, -1), "CENTER"),   # Lp
            ("ALIGN", (1, 1), (1, -1), "CENTER"),   # Data
            ("ALIGN", (2, 1), (2, -1), "CENTER"),   # Typ
            ("ALIGN", (5, 1), (5, -1), "CENTER"),   # NIP
            ("ALIGN", (6, 1), (6, -1), "RIGHT"),    # Kwota
            # przemienne tło wierszy
            *[
                ("BACKGROUND", (0, i), (-1, i), colors.HexColor("#f5f5f5"))
                for i in range(2, n - 3, 2)
            ],
            # kolorowanie KP / KW
            *[
                ("TEXTCOLOR", (2, i), (2, i),
                 colors.HexColor("#1b5e20") if data_table[i][2] == "KP"
                 else colors.HexColor("#bf360c"))
                for i in range(1, n - 3)
            ],
            # siatka
            ("GRID", (0, 0), (-1, n - 4), 0.3, colors.HexColor("#dddddd")),
            ("LINEABOVE", (0, 0), (-1, 0), 1.5, colors.HexColor("#1a237e")),
            ("LINEBELOW", (0, 0), (-1, 0), 1.5, colors.HexColor("#1a237e")),
            # podsumowanie
            ("FONTNAME", (0, n - 3), (-1, -1), "Helvetica-Bold"),
            ("FONTSIZE", (0, n - 3), (-1, -1), 8),
            ("ALIGN", (5, n - 3), (5, -1), "RIGHT"),
            ("ALIGN", (6, n - 3), (6, -1), "RIGHT"),
            ("LINEABOVE", (0, n - 3), (-1, n - 3), 1, colors.HexColor("#1a237e")),
            ("BACKGROUND", (0, n - 3), (-1, -1), colors.HexColor("#e8eaf6")),
        ])

        table = Table(data_table, colWidths=col_widths_pdf, repeatRows=1)
        table.setStyle(ts)
        story.append(table)

        # Stopka z podpisem
        story.append(Spacer(1, 1 * cm))
        sign_data = [
            ["Sporządził:", "", "Zatwierdził:"],
            ["", "", ""],
            ["_________________________", "", "_________________________"],
        ]
        sign_table = Table(sign_data, colWidths=[7 * cm, 5 * cm, 7 * cm])
        sign_table.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("TOPPADDING", (0, 1), (-1, 1), 14),
        ]))
        story.append(sign_table)

        doc.build(story)
        return out


def _fmt_dec(val: Decimal) -> str:
    return str(val.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)).replace(".", ",")
