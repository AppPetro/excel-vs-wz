import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO

# ----------------------------------------------------------------
# funkcje pomocnicze
# ----------------------------------------------------------------
def normalize_col_name(name: str) -> str:
    return name.lower().replace(" ", "").replace("\xa0", "").replace("_", "")

def clean_ean(raw: str) -> str:
    s = str(raw).strip()
    # obetnij dokładnie sufix ".0" jeśli istnieje
    if s.endswith(".0"):
        return s[:-2]
    return s

def clean_qty(raw: str) -> float:
    s = str(raw).strip()
    # usuń wszystkie białe znaki i zamień przecinek na kropkę
    s = re.sub(r"\s+", "", s).replace(",", ".")
    try:
        return float(s)
    except:
        return 0.0

# ----------------------------------------------------------------
# konfiguracja Streamlita
# ----------------------------------------------------------------
st.set_page_config(
    page_title="📋 Porównywarka Zlecenie/Zamówienie vs. WZ",
    layout="wide",
)

st.title("📋 Porównywarka Zlecenie/Zamówienie (Excel lub PDF) vs. WZ (PDF lub Excel)")
st.markdown("""
**Instrukcja:**
1. Wgraj plik Zlecenia/Zamówienia (Excel lub PDF).  
2. Wgraj plik WZ (PDF lub Excel).  
3. Aplikacja porówna ilości (EAN → ilość) i pokaże, czy wszystko się zgadza.  
""")

# -------------------------
# 1) Wgrywanie plików
# -------------------------
st.sidebar.header("Krok 1: Zlecenie/Zamówienie")
uploaded_order = st.sidebar.file_uploader(
    "Wybierz plik Zlecenia/Zamówienia (Excel lub PDF)",
    type=["xlsx", "pdf"]
)
st.sidebar.header("Krok 2: WZ")
uploaded_wz = st.sidebar.file_uploader(
    "Wybierz plik WZ (PDF lub Excel)",
    type=["pdf", "xlsx"]
)

if not uploaded_order or not uploaded_wz:
    st.info("Proszę wgrać oba pliki: Zlecenie/Zamówienie oraz WZ.")
    st.stop()

# -------------------------
# 2) Parser Excel Zlecenie/Zamówienie
# -------------------------
def parse_order_excel(f):
    df_raw = pd.read_excel(f, dtype=str, header=None)
    syn_ean = { normalize_col_name(c): c for c in [
        "Symbol","symbol","kod ean","ean","kod produktu","GTIN"
    ] }
    syn_qty = { normalize_col_name(c): c for c in [
        "Ilość","Ilosc","Quantity","Qty","sztuki",
        "ilość sztuk zamówiona","zamówiona ilość"
    ] }

    # znajdź wiersz z headerem
    header_row = None
    for idx, row in df_raw.iterrows():
        norm = [ normalize_col_name(v) for v in row.values.astype(str) ]
        if any(h in syn_ean for h in norm) and any(h in syn_qty for h in norm):
            header = row.values.tolist()
            header_row = idx
            break

    if header_row is None:
        st.error(
            "Excel Zlecenia/Zamówienia musi mieć w nagłówku kolumny EAN i Ilość.\n"
            "Sprawdziłem wszystkie wiersze i nie znalazłem."
        )
        st.stop()

    # indeksy kolumn
    col_ean_idx = next(i for i,v in enumerate(header)
                       if normalize_col_name(str(v)) in syn_ean)
    col_qty_idx = next(i for i,v in enumerate(header)
                       if normalize_col_name(str(v)) in syn_qty)

    rows = []
    for _, row in df_raw.iloc[header_row+1:].iterrows():
        ean = clean_ean(row.iloc[col_ean_idx])
        qty = clean_qty(row.iloc[col_qty_idx])
        if qty <= 0:
            continue
        rows.append([ean, qty])

    return pd.DataFrame(rows, columns=["Symbol","Ilość_Zam"])

# -------------------------
# 3) Parser PDF (oba: Zlecenie i WZ)
# -------------------------
def parse_pdf_generic(f, qty_col_name):
    rows = []
    with pdfplumber.open(f) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in text.split("\n"):
                # numer, EAN, ... , ilość, ... – taki sam wzorzec co w WZ
                m = re.match(
                    r"\s*\d+\s+(\d{13})\s+.+?\s+([\d\s]+,\d{2})\s+[\d\s]+,\d{2}$",
                    line
                )
                if not m:
                    continue
                ean = clean_ean(m.group(1))
                qty = clean_qty(m.group(2))
                if qty <= 0:
                    continue
                rows.append([ean, qty])
    df = pd.DataFrame(rows, columns=["Symbol", qty_col_name])
    return df

# -------------------------
# 4) Parser Excel WZ
# -------------------------
def parse_wz_excel(f):
    df_raw = pd.read_excel(f, dtype=str)
    syn_ean = { normalize_col_name(c): c for c in ["Kod produktu","EAN","symbol"] }
    syn_qty = { normalize_col_name(c): c for c in ["Ilość","Ilosc","Quantity","Qty"] }

    col_ean = next((c for c in df_raw.columns if normalize_col_name(c) in syn_ean), None)
    col_qty = next((c for c in df_raw.columns if normalize_col_name(c) in syn_qty), None)
    if not col_ean or not col_qty:
        st.error(
            "Excel WZ musi mieć kolumny EAN i Ilość.\n"
            f"Znalezione: {list(df_raw.columns)}"
        )
        st.stop()

    tmp = (
        df_raw[col_ean]
          .astype(str)
          .str.replace(r"\.0$", "", regex=True)
          .str.strip()
          .str.split()
          .str[-1]
    )
    rows = []
    for raw_ean, raw_qty in zip(tmp, df_raw[col_qty]):
        ean = clean_ean(raw_ean)
        qty = clean_qty(raw_qty)
        rows.append([ean, qty])
    return pd.DataFrame(rows, columns=["Symbol","Ilość_WZ"])

# -------------------------
# 5) Wybór parsera według typu pliku
# -------------------------
# Zlecenie/Zamówienie
if uploaded_order.name.lower().endswith(".xlsx"):
    df_order = parse_order_excel(uploaded_order)
else:
    df_order = parse_pdf_generic(uploaded_order, "Ilość_Zam")

# WZ
if uploaded_wz.name.lower().endswith(".xlsx"):
    df_wz = parse_wz_excel(uploaded_wz)
else:
    df_wz = parse_pdf_generic(uploaded_wz, "Ilość_WZ")

# -------------------------
# 6) Grupowanie i sumowanie
# -------------------------
df_ord_g = (
    df_order.groupby("Symbol", as_index=False)
            .agg({"Ilość_Zam": "sum"})
            .rename(columns={"Ilość_Zam": "Zamówiona_ilość"})
)
df_wz_g = (
    df_wz.groupby("Symbol", as_index=False)
         .agg({"Ilość_WZ": "sum"})
         .rename(columns={"Ilość_WZ": "Wydana_ilość"})
)

# -------------------------
# 7) Porównanie
# -------------------------
df_cmp = pd.merge(df_ord_g, df_wz_g, on="Symbol", how="outer", indicator=True)
# unikaj chained assignment – użyj przypisania
df_cmp["Zamówiona_ilość"] = df_cmp["Zamówiona_ilość"].fillna(0)
df_cmp["Wydana_ilość"]    = df_cmp["Wydana_ilość"].fillna(0)
df_cmp["Różnica"]         = df_cmp["Zamówiona_ilość"] - df_cmp["Wydana_ilość"]

def status(r):
    if r["_merge"] == "left_only":   return "Brak we WZ"
    if r["_merge"] == "right_only":  return "Brak w zamówieniu"
    return "OK" if r["Różnica"] == 0 else "Różni się"

df_cmp["Status"] = df_cmp.apply(status, axis=1)
order = ["Różni się","Brak we WZ","Brak w zamówieniu","OK"]
df_cmp["Status"] = pd.Categorical(df_cmp["Status"], categories=order, ordered=True)
df_cmp.sort_values(["Status","Symbol"], inplace=True)

# -------------------------
# 8) Wyświetlenie i eksport
# -------------------------
def highlight(row):
    color = "#c6efce" if row["Status"] == "OK" else "#ffc7ce"
    return [f"background-color: {color}" for _ in row.index]

st.markdown("### 📊 Wynik porównania")
styled = (
    df_cmp.style
          .format({"Zamówiona_ilość":"{:.0f}",
                   "Wydana_ilość":"{:.0f}",
                   "Różnica":"{:.0f}"})
          .apply(highlight, axis=1)
)
st.dataframe(styled, use_container_width=True)

def to_excel(df):
    out = BytesIO()
    writer = pd.ExcelWriter(out, engine="openpyxl")
    df.to_excel(writer, index=False, sheet_name="Porównanie")
    writer.close()
    return out.getvalue()

st.download_button(
    "⬇️ Pobierz raport Excel",
    data=to_excel(df_cmp),
    file_name="porownanie_order_vs_wz.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

all_ok = (df_cmp["Status"] == "OK").all()
if all_ok:
    st.markdown("<h4 style='color:green;'>✅ Pozycje się zgadzają</h4>", unsafe_allow_html=True)
else:
    st.markdown("<h4 style='color:red;'>❌ Pozycje się nie zgadzają</h4>", unsafe_allow_html=True)
