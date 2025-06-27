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

def find_header_and_idxs(df: pd.DataFrame, syn_ean_list: list, syn_qty_list: list):
    syn_ean_keys = {normalize_col_name(x) for x in syn_ean_list}
    syn_qty_keys = {normalize_col_name(x) for x in syn_qty_list}
    for i, row in df.iterrows():
        norm = [normalize_col_name(str(v)) for v in row.values]
        e_i = next((j for j, cell in enumerate(norm) if cell in syn_ean_keys), None)
        q_i = next((j for j, cell in enumerate(norm) if cell in syn_qty_keys), None)
        if e_i is not None and q_i is not None:
            return i, e_i, q_i
    return None, None, None

# ── Parsowanie Excela (uniwersalne) ─────────────────────────────
def parse_excel(f, syn_ean_list, syn_qty_list, col_qty_name):
    df = pd.read_excel(f, dtype=str, header=None)
    h_row, e_i, q_i = find_header_and_idxs(df, syn_ean_list, syn_qty_list)
    if h_row is None:
        st.error(f"Excel musi mieć w nagłówku kolumny EAN ({syn_ean_list}) i Ilość ({syn_qty_list}).")
        st.stop()
    rows = []
    for _, r in df.iloc[h_row+1:].iterrows():
        ean = clean_ean(r.iloc[e_i])
        qty = clean_qty(r.iloc[q_i])
        if qty > 0:
            rows.append([ean, qty])
    return pd.DataFrame(rows, columns=["Symbol", col_qty_name])

# ── Parsowanie PDF: Zlecenie/Zamówienie ─────────────────────────
ORDER_PDF_PATTERN = re.compile(r"\s*\d+\s+.+?\s+([\d\s]+,\d+)\s+\S+\s+(\d{13})")
def parse_order_pdf(f):
    rows = []
    with pdfplumber.open(f) as pdf:
        for page in pdf.pages:
            for line in (page.extract_text() or "").splitlines():
                m = ORDER_PDF_PATTERN.match(line)
                if not m:
                    continue
                qty = clean_qty(m.group(1))
                ean = clean_ean(m.group(2))
                if qty > 0:
                    rows.append([ean, qty])
    return pd.DataFrame(rows, columns=["Symbol", "Ilość_Zam"])

# ── Parsowanie PDF: WZ ───────────────────────────────────────────
def parse_wz_pdf(f):
    rows = []
    with pdfplumber.open(f) as pdf:
        for page in pdf.pages:
            for line in (page.extract_text() or "").splitlines():
                ean_m = re.search(r"\b(\d{13})\b", line)
                if not ean_m:
                    continue
                qty_matches = re.findall(r"[\d\s]+,\d{2}", line)
                if not qty_matches:
                    continue
                ean = clean_ean(ean_m.group(1))
                qty = clean_qty(qty_matches[-1])
                if qty > 0:
                    rows.append([ean, qty])
    return pd.DataFrame(rows, columns=["Symbol", "Ilość_WZ"])

# ── Streamlit UI ────────────────────────────────────────────────
st.set_page_config(page_title="📋 Porównywarka Zlecenie↔WZ", layout="wide")
st.title("📋 Porównywarka Zlecenie/Zamówienie vs. WZ")

# Instrukcja obsługi dostępna od razu
with st.expander("ℹ️ Instrukcja obsługi", expanded=True):
    st.markdown("""
**Jak to działa?**

1. W pierwszym polu **wgrywasz**:
   - **Zlecenie transportowe** (PDF)  
   - lub **Zlecenie wydania** (PDF lub Excel)
2. W drugim polu **wgrywasz**:
   - **WZ** (PDF lub Excel)

**Excel (.xlsx)**  
- Nagłówek może być w dowolnym wierszu, ale kolumny muszą mieć dokładnie nazwy (po normalizacji):  
  - **EAN**: Symbol, symbol, Kod EAN, kod ean, Kod produktu, GTIN  
  - **Ilość**: Ilość, Ilosc, Quantity, Qty, sztuki, ilość sztuk zamówiona, zamówiona ilość  
- Aplikacja usunie sufiks `.0` z EAN i konwertuje ilości w formacie `1 638,00` → `1638.00`.

**PDF – Zlecenie/Zamówienie**  
- Parsowanie według wzorca: ilość → jednostka → EAN.

**PDF – WZ**  
- Parsowanie: 13-cyfrowy EAN gdziekolwiek w linii oraz ostatni fragment `123 456,78`.

**Wynik**  
- Wyświetlana tabela z kolumnami: Symbol, Zamówiona_ilość, Wydana_ilość, Różnica, Status.  
- Zielone wiersze = OK; czerwone = rozbieżności/braki.  
- Przycisk „⬇️ Pobierz raport” pozwoli pobrać gotowy plik Excel.
""")

st.sidebar.header("Krok 1: Zlecenie/Zamówienie")
up1 = st.sidebar.file_uploader("Wybierz Zlecenie/Zamówienie", type=["xlsx","pdf"], key="order")
st.sidebar.header("Krok 2: WZ")
up2 = st.sidebar.file_uploader("Wybierz WZ", type=["xlsx","pdf"], key="wz")

if not up1 or not up2:
    st.info("Proszę wgrać oba pliki.")
    st.stop()

# Synonimy dla Excela
EAN_SYNS = ["Symbol","symbol","kod ean","ean","kod produktu","gtin"]
QTY_SYNS = ["Ilość","Ilosc","Quantity","Qty","sztuki","ilość sztuk zamówiona","zamówiona ilość"]

# Parsowanie plików
if up1.name.lower().endswith(".xlsx"):
    df1 = parse_excel(up1, EAN_SYNS, QTY_SYNS, "Ilość_Zam")
else:
    df1 = parse_order_pdf(up1)

if up2.name.lower().endswith(".xlsx"):
    df2 = parse_excel(up2, EAN_SYNS, QTY_SYNS, "Ilość_WZ")
else:
    df2 = parse_wz_pdf(up2)

# Grupowanie i porównanie
g1 = df1.groupby("Symbol", as_index=False).sum().rename(columns={"Ilość_Zam":"Zamówiona_ilość"})
g2 = df2.groupby("Symbol", as_index=False).sum().rename(columns={"Ilość_WZ":"Wydana_ilość"})
cmp = pd.merge(g1, g2, on="Symbol", how="outer", indicator=True)
cmp["Zamówiona_ilość"].fillna(0, inplace=True)
cmp["Wydana_ilość"].fillna(0, inplace=True)
cmp["Różnica"] = cmp["Zamówiona_ilość"] - cmp["Wydana_ilość"]

def status(r):
    if r["_merge"] == "left_only":
        return "Brak we WZ"
    if r["_merge"] == "right_only":
        return "Brak w zamówieniu"
    return "OK" if r["Różnica"] == 0 else "Różni się"

cmp["Status"] = cmp.apply(status, axis=1)
order = ["Różni się","Brak we WZ","Brak w zamówieniu","OK"]
cmp["Status"] = pd.Categorical(cmp["Status"], categories=order, ordered=True)
cmp.sort_values(["Status","Symbol"], inplace=True)

def highlight_row(row):
    color = "#c6efce" if row["Status"] == "OK" else "#ffc7ce"
    return [f"background-color: {color}"] * len(row)

st.markdown("### 📊 Wynik porównania")
st.dataframe(
    cmp.style
       .format({"Zamówiona_ilość":"{:.0f}", "Wydana_ilość":"{:.0f}", "Różnica":"{:.0f}"})
       .apply(highlight_row, axis=1),
    use_container_width=True
)

buf = BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as writer:
    cmp.to_excel(writer, index=False, sheet_name="Porównanie")

st.download_button("⬇️ Pobierz raport", data=buf.getvalue(), file_name="raport.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if (cmp["Status"] == "OK").all():
    st.markdown("<h4 style='color:green;'>✅ Pozycje się zgadzają</h4>", unsafe_allow_html=True)
else:
    st.markdown("<h4 style='color:red;'>❌ Pozycje się nie zgadzają</h4>", unsafe_allow_html=True)
