"""
Raport Kasowy – aplikacja Streamlit
Abacus Centrum Księgowe | Puławy
"""

import tempfile
import re
from datetime import date, datetime
from decimal import Decimal, InvalidOperation

import streamlit as st

from raport_kasowy import (
    RaportKasowy,
    KasaRecord,
    parse_xml_faktura,
    parse_bank_pdf,
    process_payroll,
)

TYPY_WYPLAT = ["wynagrodzenie", "dieta"]

# ── Konfiguracja strony ──────────────────────────────────────────────────────
st.set_page_config(
    page_title="Raport Kasowy | Abacus",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
html, body, [class*="css"] { font-family: 'Segoe UI', sans-serif; }

[data-testid="stSidebar"] {
    background-color: #f0f2f6;
    border-right: 1px solid #e0e0e0;
}
[data-testid="stSidebar"] .stMarkdown h3 {
    color: #1a237e; font-size: 0.92rem; font-weight: 700;
    margin-top: 0.8rem; margin-bottom: 0.2rem;
}

.main-header {
    background: linear-gradient(135deg, #1e3a5f 0%, #2e6da4 100%);
    color: white; padding: 2rem 2.5rem 1.6rem 2.5rem;
    border-radius: 10px; margin-bottom: 1.8rem; text-align: center;
}
.main-header h1 { font-size: 2rem; font-weight: 800; margin: 0; }
.main-header p  { font-size: 0.93rem; margin: 0.4rem 0 0 0; opacity: 0.85; }

.step-box {
    border-left: 4px solid #2e6da4; background: #f7f9fc;
    padding: 0.65rem 1rem; border-radius: 0 8px 8px 0;
    margin-bottom: 0.8rem; font-weight: 600; color: #1a237e; font-size: 0.92rem;
}
.step-box-kw {
    border-left: 4px solid #e65100; background: #fff8f5;
    padding: 0.65rem 1rem; border-radius: 0 8px 8px 0;
    margin-bottom: 0.8rem; font-weight: 600; color: #bf360c; font-size: 0.92rem;
}
.step-box-kp {
    border-left: 4px solid #2e7d32; background: #f5fbf5;
    padding: 0.65rem 1rem; border-radius: 0 8px 8px 0;
    margin-bottom: 0.8rem; font-weight: 600; color: #1b5e20; font-size: 0.92rem;
}

.metric-kp {
    background: #e8f5e9; border-left: 4px solid #2e7d32;
    border-radius: 0 8px 8px 0; padding: 0.8rem 1rem;
    font-size: 0.9rem; font-weight: 700; color: #1b5e20;
}
.metric-kw {
    background: #fff3e0; border-left: 4px solid #e65100;
    border-radius: 0 8px 8px 0; padding: 0.8rem 1rem;
    font-size: 0.9rem; font-weight: 700; color: #bf360c;
}
.metric-saldo {
    background: #e8eaf6; border-left: 4px solid #1a237e;
    border-radius: 0 8px 8px 0; padding: 0.8rem 1rem;
    font-size: 0.9rem; font-weight: 700; color: #1a237e;
}
.info-box {
    background: #e3f2fd; border-left: 4px solid #1565c0;
    border-radius: 0 6px 6px 0; padding: 0.45rem 0.8rem;
    font-size: 0.82rem; color: #0d47a1; margin-bottom: 0.5rem;
}

/* Polskie tłumaczenie file uploadera */
[data-testid="stFileUploaderDropzoneInstructions"] div span {
    visibility: hidden; font-size: 0;
}
[data-testid="stFileUploaderDropzoneInstructions"] div span::after {
    content: "Przeciągnij i upuść plik tutaj";
    visibility: visible; font-size: 1rem;
}
[data-testid="stFileUploaderDropzoneInstructions"] div small {
    visibility: hidden; font-size: 0;
}
[data-testid="stFileUploaderDropzoneInstructions"] div small::after {
    content: "Limit 200MB na plik";
    visibility: visible; font-size: 0.8rem; color: #666;
}
[data-testid="stFileUploaderDropzone"] button {
    font-size: 0 !important;
}
[data-testid="stFileUploaderDropzone"] button::after {
    content: "Przeglądaj pliki"; font-size: 0.9rem;
}

.stDownloadButton > button {
    background: linear-gradient(135deg, #1e3a5f, #2e6da4) !important;
    color: white !important; border: none !important;
    border-radius: 6px !important; font-weight: 600 !important;
}
footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)


# ── Funkcje pomocnicze ───────────────────────────────────────────────────────

def _nip_clean(v: str) -> str:
    return re.sub(r"[^0-9]", "", v or "")

def fmt_pln(val: Decimal) -> str:
    s = f"{float(val):,.2f} zł"
    return s.replace(",", " ").replace(".", ",")

def safe_decimal(txt: str) -> Decimal:
    try:
        return abs(Decimal(str(txt).replace(",", ".")))
    except InvalidOperation:
        return Decimal("0")

def records_to_df(records):
    import pandas as pd
    return pd.DataFrame([{
        "Lp.": r.lp, "Data": r.data_str(), "Typ": r.typ,
        "Nr dokumentu": r.numer_dokumentu, "Kontrahent": r.kontrahent,
        "NIP": r.nip or "", "Kwota (PLN)": float(r.kwota),
    } for r in records])

def _process_xml(uploaded, only_cash: bool, forced_typ: str,
                 raport: RaportKasowy, nip_firmy: str) -> tuple[int, str]:
    """Parsuje plik XML i dodaje rekordy do raportu. Zwraca (liczba, komunikat)."""
    with tempfile.NamedTemporaryFile(suffix=".xml", delete=False) as tmp:
        tmp.write(uploaded.read())
        tmp_path = tmp.name
    recs, diag, fmt = parse_xml_faktura(tmp_path, only_cash=only_cash,
                                         firma_nip=_nip_clean(nip_firmy))
    fmt_label = "KSeF" if fmt == "ksef" else "JPK_FA"

    if fmt == "jpk_fa" and not diag.get("has_forma_platnosci", True):
        recs, _, _ = parse_xml_faktura(tmp_path, only_cash=False,
                                        firma_nip=_nip_clean(nip_firmy))

    for r in recs:
        r.typ = forced_typ
        raport.dodaj_rekord(r)

    n = len(recs)
    skipped = diag.get("skipped", 0)
    msg = f"✅ {fmt_label}: **{n}** faktur → **{forced_typ}**"
    if skipped:
        msg += f" (pominięto {skipped} niegotówkowych)"
    return n, msg


# ── Session state ─────────────────────────────────────────────────────────────
if "records" not in st.session_state:
    st.session_state.records = []
if "payroll_rows" not in st.session_state:
    st.session_state.payroll_rows = [{"nazwisko": "", "kwota": "", "typ_wyplaty": "wynagrodzenie"}]
if "raport_okres" not in st.session_state:
    st.session_state.raport_okres = ""
if "saldo_poprzednie" not in st.session_state:
    st.session_state.saldo_poprzednie = Decimal("0")


# ════════════════════════════════════════════════════════════════════════════
#  SIDEBAR
# ════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### ⚙️ Konfiguracja")
    st.divider()

    st.markdown("### 🏢 Dane podmiotu")
    nip_firmy = st.text_input("NIP firmy", placeholder="0000000000",
                               help="Podaj NIP i kliknij 🔍 aby pobrać nazwę z rejestru MF")

    col_nip1, col_nip2 = st.columns([3, 1])
    with col_nip2:
        szukaj_nip = st.button("🔍 Szukaj", use_container_width=True, help="Pobierz nazwę z wykazu MF")

    if szukaj_nip and _nip_clean(nip_firmy):
        try:
            import requests
            from datetime import date as _date
            nip_q = _nip_clean(nip_firmy)
            url   = f"https://wl-api.mf.gov.pl/api/search/nip/{nip_q}?date={_date.today().isoformat()}"
            resp  = requests.get(url, timeout=6)
            resp.raise_for_status()
            subj  = resp.json().get("result", {}).get("subject", {})
            found = subj.get("name", "")
            if found:
                st.session_state["nazwa_z_nip"] = found
                st.success(f"✅ {found}")
            else:
                st.warning("⚠️ Nie znaleziono podmiotu o tym NIP")
        except Exception as e:
            st.error(f"❌ Błąd API MF: {e}")
    elif szukaj_nip:
        st.warning("⚠️ Podaj NIP przed wyszukaniem")

    _default_nazwa = st.session_state.get("nazwa_z_nip", "")
    nazwa_firmy = st.text_input("Nazwa firmy", value=_default_nazwa,
                                 placeholder="np. XYZ Sp. z o.o.",
                                 help="Uzupełniana automatycznie po wyszukaniu NIP")

    st.divider()
    st.markdown("### 💵 Saldo kasy")
    saldo_poprzednie_str = st.text_input(
        "Saldo końcowe poprzedniego miesiąca (PLN)",
        value="0,00",
        placeholder="0,00",
        help="Wpisz saldo końcowe kasy z poprzedniego miesiąca – "
             "staje się saldem początkowym bieżącego okresu.",
    )
    try:
        saldo_poprzednie = abs(Decimal(saldo_poprzednie_str.replace(" ", "").replace(",", ".")))
    except Exception:
        saldo_poprzednie = Decimal("0")

    st.divider()
    st.markdown("### 📅 Okres raportu")
    rok     = st.selectbox("Rok", list(range(2023, 2027))[::-1], index=1)
    miesiac = st.selectbox("Miesiąc", list(range(1, 13)),
                           format_func=lambda m: [
                               "styczeń","luty","marzec","kwiecień","maj","czerwiec",
                               "lipiec","sierpień","wrzesień","październik","listopad","grudzień"
                           ][m-1],
                           index=datetime.now().month - 1)
    okres_str = f"{rok}-{miesiac:02d}"

    st.divider()
    st.markdown("### 📂 Faktury – tryb importu")
    only_cash = st.toggle(
        "Tylko gotówkowe",
        value=True,
        help="Włączone: importuje tylko faktury z formą płatności 'gotówka'. "
             "Wyłączone: importuje wszystkie faktury z pliku (przydatne dla JPK_FA z Saldeo).",
    )

    st.divider()
    st.markdown("### 👥 Lista płac – tryb")
    collective = st.toggle(
        "Zbiorczy dokument KW",
        value=False,
        help="Włączone: jeden dokument KW dla całej listy płac. "
             "Wyłączone: osobny dokument KW dla każdego pracownika.",
    )

    st.divider()
    st.markdown(
        "<div style='font-size:0.75rem;color:#5c6bc0;text-align:center;'>"
        "🧮 Abacus Centrum Księgowe<br>Puławy · wersja 1.0</div>",
        unsafe_allow_html=True,
    )


# ════════════════════════════════════════════════════════════════════════════
#  GŁÓWNA TREŚĆ
# ════════════════════════════════════════════════════════════════════════════

st.markdown("""
<div class="main-header">
    <h1>🧾 Generator Raportu Kasowego</h1>
    <p>Automatyczne tworzenie raportu kasowego zgodnego z importem do Symfonia Finanse i Księgowość</p>
</div>
""", unsafe_allow_html=True)


# ── Sekcja 1 – Faktury kosztowe (KW) ─────────────────────────────────────────
st.markdown('<div class="step-box-kw">📤 Sekcja 1: Faktury kosztowe – zakup (gotówka wychodzi z kasy → KW)</div>',
            unsafe_allow_html=True)

c1, c2 = st.columns([3, 2])
with c1:
    jpk_kw_file = st.file_uploader(
        "Faktury kosztowe XML", type=["xml"], key="jpk_kw",
        label_visibility="collapsed",
    )
with c2:
    if jpk_kw_file:
        st.markdown('<div class="info-box">📄 Plik wczytany → wszystkie faktury jako <strong>KW</strong></div>',
                    unsafe_allow_html=True)
    else:
        st.markdown('<div class="info-box">💡 JPK_FA lub KSeF XML · np. eksport kosztów z Saldeo</div>',
                    unsafe_allow_html=True)


# ── Sekcja 2 – Faktury sprzedażowe (KP) ──────────────────────────────────────
st.markdown('<div class="step-box-kp">📥 Sekcja 2: Faktury sprzedażowe – sprzedaż (gotówka wchodzi do kasy → KP)</div>',
            unsafe_allow_html=True)

c3, c4 = st.columns([3, 2])
with c3:
    jpk_kp_file = st.file_uploader(
        "Faktury sprzedażowe XML", type=["xml"], key="jpk_kp",
        label_visibility="collapsed",
    )
with c4:
    if jpk_kp_file:
        st.markdown('<div class="info-box">📄 Plik wczytany → wszystkie faktury jako <strong>KP</strong></div>',
                    unsafe_allow_html=True)
    else:
        st.markdown('<div class="info-box">💡 JPK_FA lub KSeF XML · np. eksport sprzedaży z Saldeo</div>',
                    unsafe_allow_html=True)


# ── Sekcja 3 – Wyciąg bankowy ─────────────────────────────────────────────────
st.markdown('<div class="step-box">🏦 Sekcja 3: Wyciąg bankowy PDF – operacje gotówkowe</div>',
            unsafe_allow_html=True)

c5, c6 = st.columns([3, 2])
with c5:
    bank_file = st.file_uploader(
        "Wyciąg bankowy PDF", type=["pdf"], key="bank",
        label_visibility="collapsed",
    )
with c6:
    if bank_file:
        st.markdown('<div class="info-box">📄 Plik wczytany</div>', unsafe_allow_html=True)
    else:
        st.markdown(
            '<div class="info-box">'
            '💡 Wypłata z bankomatu → KP &nbsp;|&nbsp; Wpłata do bankomatu → KW<br>'
            'Przelew wewnętrzny / między rachunkami → wg znaku kwoty'
            '</div>',
            unsafe_allow_html=True,
        )


# ── Sekcja 4 – Lista płac i diety ────────────────────────────────────────────
st.markdown('<div class="step-box">👥 Sekcja 4: Lista płac i diety – wypłaty gotówkowe → KW</div>',
            unsafe_allow_html=True)

with st.expander("➕ Wprowadź listę płac / diet", expanded=False):
    pay_date = st.date_input(
        "Data wypłaty",
        value=date(rok, miesiac, min(28, [31,28,31,30,31,30,31,31,30,31,30,31][miesiac-1])),
    )

    col_h1, col_h2, col_h3, col_h4 = st.columns([4, 2, 2, 1])
    with col_h1: st.markdown("**Nazwisko i imię**")
    with col_h2: st.markdown("**Kwota PLN**")
    with col_h3: st.markdown("**Typ wypłaty**")

    surviving = []
    for i, row in enumerate(st.session_state.payroll_rows):
        col_n, col_k, col_t, col_d = st.columns([4, 2, 2, 1])
        with col_n:
            naz = st.text_input(f"naz_{i}", value=row.get("nazwisko",""), key=f"naz_{i}",
                                label_visibility="collapsed",
                                placeholder=f"Nazwisko i imię #{i+1}")
        with col_k:
            kwt = st.text_input(f"kwt_{i}", value=row.get("kwota",""), key=f"kwt_{i}",
                                label_visibility="collapsed", placeholder="Kwota PLN")
        with col_t:
            tw = row.get("typ_wyplaty", "wynagrodzenie")
            tw_idx = TYPY_WYPLAT.index(tw) if tw in TYPY_WYPLAT else 0
            typ_w = st.selectbox(
                f"typ_{i}", TYPY_WYPLAT, index=tw_idx, key=f"typ_{i}",
                label_visibility="collapsed",
                format_func=lambda x: "💰 Wynagrodzenie" if x == "wynagrodzenie" else "✈️ Dieta",
            )
        with col_d:
            delete = st.button("🗑", key=f"del_{i}", help="Usuń wiersz")
        if not delete or len(st.session_state.payroll_rows) == 1:
            surviving.append({"nazwisko": naz, "kwota": kwt, "typ_wyplaty": typ_w})

    if len(surviving) != len(st.session_state.payroll_rows):
        st.session_state.payroll_rows = surviving
        st.rerun()

    if st.button("➕ Dodaj wiersz"):
        st.session_state.payroll_rows.append(
            {"nazwisko": "", "kwota": "", "typ_wyplaty": "wynagrodzenie"}
        )
        st.rerun()

    total_plac = sum(safe_decimal(r["kwota"]) for r in surviving if r.get("kwota",""))
    if total_plac > 0:
        n_w = sum(1 for r in surviving if r.get("typ_wyplaty","wynagrodzenie") == "wynagrodzenie" and r.get("kwota",""))
        n_d = sum(1 for r in surviving if r.get("typ_wyplaty","") == "dieta" and r.get("kwota",""))
        parts = []
        if n_w: parts.append(f"{n_w} wynagrodzeń")
        if n_d: parts.append(f"{n_d} diet")
        st.markdown(
            f'<div class="info-box">💰 Łącznie: <strong>{fmt_pln(total_plac)}</strong>'
            f' ({", ".join(parts)})</div>',
            unsafe_allow_html=True,
        )


# ── Przycisk Generuj ──────────────────────────────────────────────────────────
st.markdown("---")
_, col_mid, _ = st.columns([2, 1, 2])
with col_mid:
    process = st.button("⚙️ Generuj raport", type="primary", use_container_width=True)

if process:
    raport  = RaportKasowy(okres=okres_str)
    errors  = []
    summary = []

    # Faktury kosztowe → KW
    if jpk_kw_file:
        try:
            n, msg = _process_xml(jpk_kw_file, only_cash, "KW", raport, nip_firmy)
            summary.append(msg)
            if n == 0:
                st.warning("⚠️ Faktury kosztowe: brak wyników. Spróbuj wyłączyć opcję "
                           "**Tylko gotówkowe** w panelu bocznym.")
        except Exception as e:
            errors.append(f"Faktury kosztowe: {e}")

    # Faktury sprzedażowe → KP
    if jpk_kp_file:
        try:
            n, msg = _process_xml(jpk_kp_file, only_cash, "KP", raport, nip_firmy)
            summary.append(msg)
            if n == 0:
                st.warning("⚠️ Faktury sprzedażowe: brak wyników. Spróbuj wyłączyć opcję "
                           "**Tylko gotówkowe** w panelu bocznym.")
        except Exception as e:
            errors.append(f"Faktury sprzedażowe: {e}")

    # Wyciąg bankowy
    if bank_file:
        try:
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                tmp.write(bank_file.read())
                tmp_path = tmp.name
            recs = parse_bank_pdf(tmp_path)
            for r in recs:
                raport.dodaj_rekord(r)
            n_kp = sum(1 for r in recs if r.typ == "KP")
            n_kw = sum(1 for r in recs if r.typ == "KW")
            summary.append(f"✅ Wyciąg bankowy: **{len(recs)}** operacji "
                            f"(KP: {n_kp}, KW: {n_kw})")
        except Exception as e:
            errors.append(f"Wyciąg bankowy: {e}")

    # Lista płac / diety → KW
    valid_rows = [r for r in st.session_state.payroll_rows
                  if r.get("nazwisko","").strip() and r.get("kwota","").strip()]
    if valid_rows:
        try:
            recs = process_payroll(valid_rows, pay_date, collective=collective)
            for r in recs:
                raport.dodaj_rekord(r)
            n_w = sum(1 for r in recs if "Dieta" not in r.kontrahent)
            n_d = sum(1 for r in recs if "Dieta" in r.kontrahent)
            parts = []
            if n_w: parts.append(f"{n_w} wynagrodzeń")
            if n_d: parts.append(f"{n_d} diet")
            summary.append(f"✅ Lista płac: **{len(recs)}** dokumentów KW "
                            f"({', '.join(parts)})")
        except Exception as e:
            errors.append(f"Lista płac: {e}")

    for msg in summary:
        st.success(msg)
    for err in errors:
        st.error(f"❌ {err}")

    if not raport._records:
        st.warning("⚠️ Brak danych – wgraj co najmniej jedno źródło danych.")
    else:
        st.session_state.records           = raport._prepare()
        st.session_state.raport_okres      = okres_str
        st.session_state.saldo_poprzednie  = saldo_poprzednie
        st.rerun()


# ── Podgląd i pobieranie ──────────────────────────────────────────────────────
if st.session_state.records:
    records = st.session_state.records

    st.markdown("---")
    st.markdown('<div class="step-box">📊 Podgląd i pobieranie raportu</div>',
                unsafe_allow_html=True)

    total_kp       = sum(r.kwota for r in records if r.typ == "KP")
    total_kw       = sum(r.kwota for r in records if r.typ == "KW")
    saldo_pocz     = st.session_state.saldo_poprzednie
    saldo_konc     = saldo_pocz + total_kp - total_kw
    n_kp           = sum(1 for r in records if r.typ == "KP")
    n_kw           = sum(1 for r in records if r.typ == "KW")

    m1, m2, m3, m4, m5, m6 = st.columns(6)
    with m1:
        st.markdown(
            f'<div class="metric-saldo">🏦 Saldo początkowe<br>'
            f'<span style="font-size:1.1rem">{fmt_pln(saldo_pocz)}</span><br>'
            f'<small>stan na początek okresu</small></div>', unsafe_allow_html=True)
    with m2:
        st.markdown(
            f'<div class="metric-kp">📥 Przychody KP<br>'
            f'<span style="font-size:1.1rem">{fmt_pln(total_kp)}</span><br>'
            f'<small>{n_kp} dokumentów</small></div>', unsafe_allow_html=True)
    with m3:
        st.markdown(
            f'<div class="metric-kw">📤 Rozchody KW<br>'
            f'<span style="font-size:1.1rem">{fmt_pln(total_kw)}</span><br>'
            f'<small>{n_kw} dokumentów</small></div>', unsafe_allow_html=True)
    with m4:
        kolor_konc = "#1b5e20" if saldo_konc >= 0 else "#b71c1c"
        st.markdown(
            f'<div class="metric-saldo">💰 Saldo końcowe<br>'
            f'<span style="font-size:1.1rem;color:{kolor_konc}">{fmt_pln(saldo_konc)}</span><br>'
            f'<small>stan na koniec okresu</small></div>', unsafe_allow_html=True)
    with m5:
        st.markdown(
            f'<div class="metric-saldo">📋 Łącznie<br>'
            f'<span style="font-size:1.1rem">{len(records)}</span><br>'
            f'<small>dokumentów kasowych</small></div>', unsafe_allow_html=True)
    with m6:
        st.markdown(
            f'<div class="info-box" style="padding:0.8rem">'
            f'<strong>Okres:</strong> {st.session_state.raport_okres}<br>'
            f'<strong>Następny mies.:</strong><br>saldo pocz. = {fmt_pln(saldo_konc)}</div>',
            unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    import pandas as pd
    df = records_to_df(records)
    df_disp = df.copy()
    df_disp["Kwota (PLN)"] = df_disp["Kwota (PLN)"].apply(
        lambda x: f"{x:,.2f}".replace(",", "\u00a0").replace(".", ",")
    )

    def highlight_typ(row):
        c = "#e8f5e9" if row["Typ"] == "KP" else "#fff3e0"
        return [f"background-color: {c}"] * len(row)

    st.dataframe(
        df_disp.style.apply(highlight_typ, axis=1),
        use_container_width=True, hide_index=True, height=370,
    )

    raport_dl = RaportKasowy(okres=st.session_state.raport_okres)
    raport_dl._records      = list(records)
    raport_dl.saldo_pocz    = st.session_state.saldo_poprzednie
    raport_dl.saldo_konc    = (st.session_state.saldo_poprzednie
                               + sum(r.kwota for r in records if r.typ == "KP")
                               - sum(r.kwota for r in records if r.typ == "KW"))

    st.markdown("<br>", unsafe_allow_html=True)
    dl1, dl2, dl3 = st.columns([3, 3, 1])

    with dl1:
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            xlsx_path = raport_dl.eksportuj_xlsx(tmp.name)
        fn_xlsx = f"raport_kasowy_{st.session_state.raport_okres.replace('-','_')}.xlsx"
        st.download_button(
            "⬇️ Pobierz XLSX (import Symfonia)",
            data=open(xlsx_path, "rb").read(), file_name=fn_xlsx,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with dl2:
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            pdf_path = raport_dl.eksportuj_pdf(tmp.name)
        fn_pdf = f"raport_kasowy_{st.session_state.raport_okres.replace('-','_')}.pdf"
        st.download_button(
            "⬇️ Pobierz PDF (wydruk / archiwum)",
            data=open(pdf_path, "rb").read(), file_name=fn_pdf,
            mime="application/pdf",
            use_container_width=True,
        )

    with dl3:
        if st.button("🗑️ Wyczyść", use_container_width=True, help="Zacznij od nowa"):
            st.session_state.records = []
            st.session_state.payroll_rows = [{"nazwisko": "", "kwota": "", "typ_wyplaty": "wynagrodzenie"}]
            st.rerun()

    st.markdown("""
    <div class="info-box" style="margin-top:1rem">
    📌 <strong>Import do Symfonii F&amp;K:</strong> Moduł Kasa → Import danych →
    wskaż plik XLSX. Kolumny: Lp. | Data | Typ (KP/KW) | Nr dokumentu | Kontrahent | NIP | Kwota
    </div>
    """, unsafe_allow_html=True)
