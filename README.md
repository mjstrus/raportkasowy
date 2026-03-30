# 🧾 Raport Kasowy

Generator raportu kasowego zgodny z importem do **Symfonia Finanse i Księgowość**.

## Funkcje

- 📂 **JPK_FA (XML)** – wyodrębnia faktury opłacone gotówką → dokumenty KP
- 🏦 **Wyciąg bankowy (PDF)** – rozpoznaje operacje bankomatowe:
  - Wypłaty z bankomatu → KP (zasilenie kasy)
  - Wpłaty do bankomatu → KW (odprowadzenie do banku)
- 👥 **Lista płac** – generuje KW dla każdego pracownika lub zbiorczy
- 📊 Podgląd tabeli z kolorami KP/KW
- ⬇️ Eksport do **XLSX** (import Symfonia) i **PDF** (wydruk / archiwum)

## Numeracja dokumentów

Format: `KP/001/01/2025` lub `KW/001/01/2025`  
(typ / nr porządkowy / miesiąc / rok – osobny licznik dla KP i KW)

## Uruchomienie lokalne

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Wdrożenie na Streamlit Cloud

1. Utwórz repozytorium na GitHub i wgraj pliki
2. Wejdź na [share.streamlit.io](https://share.streamlit.io)
3. New app → wskaż repo → `app.py` → Deploy

## Pliki

| Plik | Opis |
|------|------|
| `app.py` | Aplikacja Streamlit (interfejs) |
| `raport_kasowy.py` | Moduł logiki – parsery i generatory |
| `requirements.txt` | Zależności Python |

---
*Abacus Centrum Księgowe · Puławy*
