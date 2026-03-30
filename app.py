"""
Raport Kasowy – aplikacja Streamlit
Abacus Centrum Księgowe | Puławy
"""

import tempfile
from datetime import date, datetime
from decimal import Decimal, InvalidOperation

import streamlit as st

from raport_kasowy import (
    RaportKasowy,
    KasaRecord,
    parse_jpk_fa,
    parse_bank_pdf,
    process_payroll,
)

# ── Konfiguracja strony ──────────────────────────────────────────────────────
st.set_page_config(
    page_title="Raport Kasowy | Abacus",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS – styl jak w Informacji Dodatkowej ───────────────────────────────────
st.markdown("""
<style>
html, body, [class*="css"] { font-family: 'Segoe UI', sans-serif; }

/* Sidebar */
[data-testid="stSidebar"] {
    background-color: #f0f2f6;
    border-right: 1px solid #e0e0e0;
}
[data-testid="stSidebar"] .stMarkdown h3 {
    color: #1a237e;
    font-size: 0.92rem;
    font-weight: 700;
    margin-top: 0.8rem;
    margin-bottom: 0.2rem;
}

/* Główny nagłówek */
.main-header {
    background: linear-gradient(135deg, #1e3a5f 0%, #2e6da4 100%);
    color: white;
    padding: 2rem 2.5rem 1.6rem 2.5rem;
    border-radius: 10px;
    margin-bottom: 1.8rem;
    text-align: center;
}
.main-header h1 { font-size: 2rem; font-weight: 800; margin: 0; }
.main-header p  { font-size: 0.93rem; margin: 0.4rem 0 0 0; opacity: 0.85; }

/* Kroki */
.step-box {
    border-left: 4px solid #2e6da4;
    background: #f7f9fc;
    padding: 0.65rem 1rem;
    border-radius: 0 8px 8px 0;
    margin-bottom: 1rem;
    font-weight: 600;
    color: #1a237e;
    font-size: 0.92rem;
}

/* Metryki */
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

/* Info boxy */
.info-box {
    background: #e3f2fd; border-left: 4px solid #1565c0;
    border-radius: 0 6px 6px 0; padding: 0.45rem 0.8rem;
    font-size: 0.82rem; color: #0d47a1; margin-bottom: 0.5rem;
}

/* ── Polskie tłumaczenie file uploadera ── */

/* "Drag and drop files here" */
[data-testid="stFileUploaderDropzoneInstructions"] div span {
    visibility: hidden;
    font-size: 0;
}
[data-testid="stFileUploaderDropzoneInstructions"] div span::after {
    content: "Przeciągnij i upuść pliki tutaj";
    visibility: visible;
    font-size: 1rem;
}

/* "Limit 200MB per file • PDF, DOCX" itp. */
[data-testid="stFileUploaderDropzoneInstructions"] div small {
    visibility: hidden;
    font-size: 0;
}
[data-testid="stFileUploaderDropzoneInstructions"] div small::after {
    content: "Limit 200MB na plik";
    visibility: visible;
    font-size: 0.8rem;
    color: #666;
}

/* Przycisk "Browse files" */
[data-testid="stFileUploaderDropzone"] button {
    font-size: 0 !important;
}
[data-testid="stFileUploaderDropzone"] button::after {
    content: "Przeglądaj pliki";
    font-size: 0.9rem;
}

/* Komunikat po wgraniu: "Drag and drop file here" (singular) */
[data-testid="stFileUploader"] section div small {
    visibility: hidden;
    font-size: 0;
}
[data-testid="stFileUploader"] section div small::after {
    content: "Limit 200MB na plik";
    visibility: visible;
    font-size: 0.8rem;
    color: #666;
}

/* Przyciski pobierania */
.stDownloadButton > button {
    background: linear-gradient(135deg, #1e3a5f, #2e6da4) !important;
    color: white !important; border: none !important;
    border-radius: 6px !important; font-weight: 600 !important;
}
footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)


# ── Funkcje pomocnicze ───────────────────────────────────────────────────────

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
    rows = []
    for r in records:
        rows.append({
            "Lp.": r.lp,
            "Data": r.data_str(),
            "Typ": r.typ,
            "Nr dokumentu": r.numer_dokumentu,
            "Kontrahent": r.kontrahent,
            "NIP": r.nip or "",
            "Kwota (PLN)": float(r.kwota),
        })
    return pd.DataFrame(rows)


# ── Session state ─────────────────────────────────────────────────────────────
if "records" not in st.session_state:
    st.session_state.records = []
if "payroll_rows" not in st.session_state:
    st.session_state.payroll_rows = [{"nazwisko": "", "kwota": ""}]
if "raport_okres" not in st.session_state:
    st.session_state.raport_okres = ""


# ════════════════════════════════════════════════════════════════════════════
#  SIDEBAR
# ════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### ⚙️ Konfiguracja")
    st.divider()

    st.markdown("### 🏢 Dane podmiotu")
    nazwa_firmy = st.text_input("Nazwa firmy", placeholder="np. XYZ Sp. z o.o.")
    nip_firmy   = st.text_input("NIP firmy",   placeholder="0000000000")

    st.divider()
    st.markdown("### 📅 Okres raportu")
    rok     = st.selectbox("Rok",    list(range(2023, 2027))[::-1], index=1)
    miesiac = st.selectbox("Miesiąc", list(range(1, 13)),
                           format_func=lambda m: datetime(2000, m, 1).strftime("%B"),
                           index=datetime.now().month - 1)
    okres_str = f"{rok}-{miesiac:02d}"

    st.divider()
    st.markdown("### 👥 Lista płac – tryb")
    collective = st.toggle(
        "Zbiorczy dokument KW",
        value=False,
        help="Włącz = jeden dokument KW dla całej listy płac",
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


# ── Krok 1 – JPK_FA ──────────────────────────────────────────────────────────
st.markdown('<div class="step-box">📂 Krok 1: Wgraj JPK_FA (XML) – faktury opłacone gotówką → KP</div>',
            unsafe_allow_html=True)

c1, c2 = st.columns([3, 2])
with c1:
    jpk_file = st.file_uploader("JPK_FA", type=["xml"], key="jpk",
                                 label_visibility="collapsed")
with c2:
    if jpk_file:
        st.markdown('<div class="info-box">📄 Plik wczytany – kliknij „Przetwórz źródła"</div>',
                    unsafe_allow_html=True)
    else:
        st.markdown('<div class="info-box">💡 Obsługiwane: JPK_FA v3 i v4 (XML)</div>',
                    unsafe_allow_html=True)


# ── Krok 2 – Wyciąg bankowy ──────────────────────────────────────────────────
st.markdown('<div class="step-box">🏦 Krok 2: Wgraj wyciąg bankowy (PDF) – operacje bankomatowe</div>',
            unsafe_allow_html=True)

c3, c4 = st.columns([3, 2])
with c3:
    bank_file = st.file_uploader("Wyciąg bankowy PDF", type=["pdf"], key="bank",
                                  label_visibility="collapsed")
with c4:
    if bank_file:
        st.markdown('<div class="info-box">📄 Wyciąg wczytany</div>', unsafe_allow_html=True)
    else:
        st.markdown(
            '<div class="info-box">💡 Wypłaty z bankomatu → KP &nbsp;|&nbsp; Wpłaty do bankomatu → KW</div>',
            unsafe_allow_html=True,
        )


# ── Krok 3 – Lista płac ──────────────────────────────────────────────────────
st.markdown('<div class="step-box">👥 Krok 3: Lista płac – wypłaty gotówkowe pracownikom → KW</div>',
            unsafe_allow_html=True)

with st.expander("➕ Wprowadź listę płac", expanded=False):
    pay_date = st.date_input(
        "Data wypłaty",
        value=date(rok, miesiac, min(28, [31,28,31,30,31,30,31,31,30,31,30,31][miesiac-1])),
    )
    st.markdown("**Pracownicy i kwoty:**")

    surviving = []
    for i, row in enumerate(st.session_state.payroll_rows):
        col_n, col_k, col_d = st.columns([4, 2, 1])
        with col_n:
            naz = st.text_input(f"naz_{i}", value=row["nazwisko"], key=f"naz_{i}",
                                label_visibility="collapsed", placeholder=f"Nazwisko i imię #{i+1}")
        with col_k:
            kwt = st.text_input(f"kwt_{i}", value=row["kwota"], key=f"kwt_{i}",
                                label_visibility="collapsed", placeholder="Kwota PLN")
        with col_d:
            delete = st.button("🗑", key=f"del_{i}", help="Usuń wiersz")
        if not delete or len(st.session_state.payroll_rows) == 1:
            surviving.append({"nazwisko": naz, "kwota": kwt})

    if len(surviving) != len(st.session_state.payroll_rows):
        st.session_state.payroll_rows = surviving
        st.rerun()

    if st.button("➕ Dodaj pracownika"):
        st.session_state.payroll_rows.append({"nazwisko": "", "kwota": ""})
        st.rerun()

    total_plac = sum(safe_decimal(r["kwota"]) for r in surviving if r["kwota"])
    if total_plac > 0:
        st.markdown(
            f'<div class="info-box">💰 Suma listy płac: <strong>{fmt_pln(total_plac)}</strong></div>',
            unsafe_allow_html=True,
        )


# ── Przycisk Przetwórz ────────────────────────────────────────────────────────
st.markdown("---")
_, col_mid, _ = st.columns([2, 1, 2])
with col_mid:
    process = st.button("⚙️ Przetwórz źródła", type="primary", use_container_width=True)

if process:
    raport  = RaportKasowy(okres=okres_str)
    errors  = []
    summary = []

    # JPK_FA
    if jpk_file:
        try:
            with tempfile.NamedTemporaryFile(suffix=".xml", delete=False) as tmp:
                tmp.write(jpk_file.read())
                tmp_path = tmp.name
            recs = parse_jpk_fa(tmp_path)
            for r in recs:
                raport.dodaj_rekord(r)
            summary.append(f"✅ JPK_FA: **{len(recs)}** faktur gotówkowych")
        except Exception as e:
            errors.append(f"JPK_FA: {e}")

    # Wyciąg
    if bank_file:
        try:
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                tmp.write(bank_file.read())
                tmp_path = tmp.name
            recs = parse_bank_pdf(tmp_path)
            for r in recs:
                raport.dodaj_rekord(r)
            summary.append(f"✅ Wyciąg bankowy: **{len(recs)}** operacji bankomatowych")
        except Exception as e:
            errors.append(f"Wyciąg bankowy: {e}")

    # Lista płac
    valid_rows = [r for r in st.session_state.payroll_rows
                  if r["nazwisko"].strip() and r["kwota"].strip()]
    if valid_rows:
        try:
            recs = process_payroll(valid_rows, pay_date, collective=collective)
            for r in recs:
                raport.dodaj_rekord(r)
            summary.append(f"✅ Lista płac: **{len(recs)}** dokumentów KW")
        except Exception as e:
            errors.append(f"Lista płac: {e}")

    for msg in summary:
        st.success(msg)
    for err in errors:
        st.error(f"❌ {err}")

    if not raport._records:
        st.warning("⚠️ Brak danych do przetworzenia – wgraj co najmniej jedno źródło.")
    else:
        st.session_state.records     = raport._prepare()
        st.session_state.raport_okres = okres_str
        st.rerun()


# ── Podgląd i pobieranie ──────────────────────────────────────────────────────
if st.session_state.records:
    records = st.session_state.records

    st.markdown("---")
    st.markdown('<div class="step-box">📊 Krok 4: Podgląd i pobieranie raportu</div>',
                unsafe_allow_html=True)

    total_kp = sum(r.kwota for r in records if r.typ == "KP")
    total_kw = sum(r.kwota for r in records if r.typ == "KW")
    saldo    = total_kp - total_kw
    n_kp     = sum(1 for r in records if r.typ == "KP")
    n_kw     = sum(1 for r in records if r.typ == "KW")

    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.markdown(
            f'<div class="metric-kp">📥 Przychody KP<br>'
            f'<span style="font-size:1.25rem">{fmt_pln(total_kp)}</span><br>'
            f'<small>{n_kp} dokumentów</small></div>', unsafe_allow_html=True)
    with m2:
        st.markdown(
            f'<div class="metric-kw">📤 Rozchody KW<br>'
            f'<span style="font-size:1.25rem">{fmt_pln(total_kw)}</span><br>'
            f'<small>{n_kw} dokumentów</small></div>', unsafe_allow_html=True)
    with m3:
        kolor = "#1b5e20" if saldo >= 0 else "#b71c1c"
        st.markdown(
            f'<div class="metric-saldo">💰 Saldo kasy<br>'
            f'<span style="font-size:1.25rem;color:{kolor}">{fmt_pln(saldo)}</span><br>'
            f'<small>Okres: {st.session_state.raport_okres}</small></div>',
            unsafe_allow_html=True)
    with m4:
        st.markdown(
            f'<div class="metric-saldo">📋 Łącznie<br>'
            f'<span style="font-size:1.25rem">{len(records)}</span><br>'
            f'<small>dokumentów kasowych</small></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Tabela
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
        use_container_width=True,
        hide_index=True,
        height=370,
    )

    # Generowanie plików do pobrania
    raport_dl = RaportKasowy(okres=st.session_state.raport_okres)
    raport_dl._records = list(records)

    st.markdown("<br>", unsafe_allow_html=True)
    dl1, dl2, dl3 = st.columns([3, 3, 1])

    with dl1:
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            xlsx_path = raport_dl.eksportuj_xlsx(tmp.name)
        xlsx_bytes = open(xlsx_path, "rb").read()
        fn_xlsx = f"raport_kasowy_{st.session_state.raport_okres.replace('-','_')}.xlsx"
        st.download_button(
            "⬇️ Pobierz XLSX (import Symfonia)",
            data=xlsx_bytes, file_name=fn_xlsx,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with dl2:
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            pdf_path = raport_dl.eksportuj_pdf(tmp.name)
        pdf_bytes = open(pdf_path, "rb").read()
        fn_pdf = f"raport_kasowy_{st.session_state.raport_okres.replace('-','_')}.pdf"
        st.download_button(
            "⬇️ Pobierz PDF (wydruk / archiwum)",
            data=pdf_bytes, file_name=fn_pdf,
            mime="application/pdf",
            use_container_width=True,
        )

    with dl3:
        if st.button("🗑️ Wyczyść", use_container_width=True, help="Zacznij od nowa"):
            st.session_state.records = []
            st.session_state.payroll_rows = [{"nazwisko": "", "kwota": ""}]
            st.rerun()

    st.markdown("""
    <div class="info-box" style="margin-top:1rem">
    📌 <strong>Import do Symfonii F&amp;K:</strong> Moduł Kasa → Import danych →
    wskaż plik XLSX. Kolumny: Lp. | Data | Typ (KP/KW) | Nr dokumentu | Kontrahent | NIP | Kwota
    </div>
    """, unsafe_allow_html=True)
