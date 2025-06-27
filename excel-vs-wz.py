import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO

# ── Funkcje pomocnicze ──────────────────────────────────────────
def normalize_col_name(name: str) -> str:
    return name.lower().replace(" ", "").replace("\xa0", "").replace("_", "")

def clean_ean(raw: str) -> str:
    s = str(raw).strip()
    return s[:-2] if s.endswith(".0") else s

def clean_qty(raw: str) -> float:
    s = re.sub(r"\s+", "", str(raw)).replace(",", ".")
    try:
        return float(s)
    except:
        return 0.0

def find_header_and_idxs(df: pd.DataFrame, syn_ean: dict, syn_qty: dict):
    for i, row in df.iterrows():
        norm = [normalize_col_name(v) for v in row.values.astype(str)]
        if any(k in syn_ean for k in norm) and any(k in syn_qty for k in norm):
            e_i = next(idx for idx, val in enumerate(norm) if val in syn_ean)
            q_i = next(idx for idx, val in enumerate(norm) if val in syn_qty)
            return i, e_i, q_i
    return None, None, None

# ── Parsowanie Excela ───────────────────────────────────────────
def parse_excel(f, syn_ean_list, syn_qty_list, col_qty_name):
    df = pd.read_excel(f, dtype=str, header=None)
    syn_ean = {normalize_col_name(x): x for x in syn_ean_list}
    syn_qty = {normalize_col_name(x): x for x in syn_qty_list}
    h_row, e_i, q_i = find_header_and_idxs(df, syn_ean, syn_qty)
    if h_row is None:
        st.error(f"Excel musi mieć nagłówek EAN {syn_ean_list} i Ilość {syn_qty_list}.")
        st.stop()
    out = []
    for _, r in df.iloc[h_row+1:].iterrows():
        ean = clean_ean(r.iloc[e_i])
        qty = clean_qty(r.iloc[q_i])
        if qty > 0:
            out.append([ean, qty])
    return pd.DataFrame(out, columns=["Symbol", col_qty_name])

# ── Parsowanie PDF ──────────────────────────────────────────────
PDF_PATTERN = r"\s*\d+\s+(\d{13})\s+.+?\s+([\d\s]+,\d{2})\s+[\d\s]+,\d{2}$"
def parse_pdf(f, col_qty_name):
    rows = []
    with pdfplumber.open(f) as pdf:
        for p in pdf.pages:
            for line in (p.extract_text() or "").splitlines():
                m = re.match(PDF_PATTERN, line)
                if not m:
                    continue
                ean = clean_ean(m.group(1))
                qty = clean_qty(m.group(2))
                if qty > 0:
                    rows.append([ean, qty])
    return pd.DataFrame(rows, columns=["Symbol", col_qty_name])

# ── UI ──────────────────────────────────────────────────────────
st.set_page_config(page_title="📋 Porównywarka Zlecenie↔WZ", layout="wide")
st.title("📋 Porównywarka Zlecenie/Zamówienie vs. WZ")

# ── Instrukcja obsługi (dostępna od razu) ───────────────────────
with st.expander("ℹ️ Instrukcja obsługi", expanded=True):
    st.markdown('''
**Jak to działa?**  
- Wgrywasz dwa pliki: Zlecenie/Zamówienie (pierwszy uploader) i WZ (drugi uploader).  
- Oba mogą być w formacie **Excel (.xlsx)** lub **PDF**, niezależnie od siebie.

**Dla Excela (.xlsx):**  
1. Aplikacja sama wyszukuje wiersz nagłówka (może być w dowolnej linii).  
2. Rozpoznaje kolumnę z kodami **EAN** i kolumnę z **ilościami** wg poniższych synonimów:  
   - **EAN**: Symbol, symbol, Kod EAN, kod ean, Kod produktu, GTIN  
   - **Ilość**: Ilość, Ilosc, Quantity, Qty, sztuki, ilość sztuk zamówiona, zamówiona ilość  
3. Usuwa z EAN ewentualny sufiks `.0` (np. `4250231542008.0` → `4250231542008`).  
4. Ilości w formacie `1 638,00` lub `1638,00` poprawnie konwertuje (usuwa spacje, zamienia przecinek na kropkę).

**Dla PDF:**  
- Aplikacja skanuje każdą linijkę tekstu i wyciąga EAN oraz ilość zgodnie z wzorcem:
