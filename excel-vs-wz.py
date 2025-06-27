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
    if s.endswith(".0"):
        return s[:-2]
    return s

def clean_qty(raw: str) -> float:
    s = str(raw).strip()
    s = re.sub(r"\s+", "", s).replace(",", ".")
    try:
        return float(s)
    except:
        return 0.0

def find_header_and_idxs(df_raw: pd.DataFrame, syn_ean: dict, syn_qty: dict):
    """Zwraca (header_row_index, ean_col_idx, qty_col_idx) lub (None, None, None)."""
    for idx, row in df_raw.iterrows():
        norm = [normalize_col_name(v) for v in row.values.astype(str)]
        if any(h in syn_ean for h in norm) and any(h in syn_qty for h in norm):
            ean_idx = next(i for i, v in enumerate(norm) if v in syn_ean)
            qty_idx = next(i for i, v in enumerate(norm) if v in syn_qty)
            return idx, ean_idx, qty_idx
    return None, None, None

# ----------------------------------------------------------------
# parsowanie Excela (wspólne dla obu uploaderów)
# ----------------------------------------------------------------
def parse_excel_generic(f, syn_ean_list, syn_qty_list, col_name_qty, col_name_ean="Symbol"):
    df_raw = pd.read_excel(f, dtype=str, header=None)
    syn_ean = {normalize_col_name(c): c for c in syn_ean_list}
    syn_qty = {normalize_col_name(c): c for c in syn_qty_list}

    header_row, ean_idx, qty_idx = find_header_and_idxs(df_raw, syn_ean, syn_qty)
    if header_row is None:
        st.error(
            f"Excel musi mieć w nagłówku kolumny EAN ({syn_ean_list}) i Ilość ({syn_qty_list}).\n"
            "Nie znalazłem ich w żadnym wierszu."
        )
        st.stop()

    rows = []
    for _, row in df_raw.iloc[header_row + 1 :].iterrows():
        ean = clean_ean(row.iloc[ean_idx])
        qty = clean_qty(row.iloc[qty_idx])
        if qty <= 0:
            continue
        rows.append([ean, qty])

    return pd.DataFrame(rows, columns=[col_name_ean, col_name_qty])

# ----------------------------------------------------------------
# parsowanie PDF (wspólne dla obu uploaderów)
# ----------------------------------------------------------------
def parse_pdf_generic(f, pattern, col_name_qty, col_name_ean="Symbol"):
    rows = []
    with pdfplumber.open(f) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in text.split("\n"):
                m = re.match(pattern, line)
                if not m:
                    continue
                ean = clean_ean(m.group(1))
                qty = clean_qty(m.group(2))
                if qty <= 0:
                    continue
                rows.append([ean, qty])
    return pd.DataFrame(rows, columns=[col_name_ean, col_name_qty])

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

st.sidebar.header("Krok 1: Zlecenie/Zamówienie")
uploaded_order = st.sidebar.file_uploader("Wybierz Zlecenie/Zamówienie", type=["xlsx","pdf"])
st.sidebar.header("Krok 2: WZ")
uploaded_wz     = st.sidebar.file_uploader("Wybierz WZ",               type=["pdf","xlsx"])

if not uploaded_order or not uploaded_wz:
    st.info("Proszę wgrać oba pliki.")
    st.stop()

# ------------------------------------
# 1) Parsowanie Zlecenia/Zamówienia
# ------------------------------------
syn_ean_all = ["Symbol","symbol","kod ean","ean","kod produktu","GTIN"]
syn_qty_all = ["Ilość","Ilosc","Quantity","Qty","sztuki","ilość sztuk zamówiona","zamówiona ilość"]

if uploaded_order.name.lower().endswith(".xlsx"):
    df_order = parse_excel_generic(
        uploaded_order,
        syn_ean_list=syn_ean_all,
        syn_qty_list=syn_qty_all,
        col_name_qty="Ilość_Zam"
    )
else:
    # regex: numer, EAN, ... , ilość, jednostka, waga
    pattern_order = r"\s*\d+\s+(\d{13})\s+.+?\s+([\d\s]+,\d{2})\s+\S+\s+[\d\s]+,\d{2}$"
    df_order = parse_pdf_generic(
        uploaded_order,
        pattern=pattern_order,
        col_name_qty="Ilość_Zam"
    )

# -------------------------
# 2) Parsowanie WZ
# -------------------------
if uploaded_wz.name.lower().endswith(".xlsx"):
    df_wz = parse_excel_generic(
        uploaded_wz,
        syn_ean_list=syn_ean_all,     # TE SAME synonimy co wyżej
        syn_qty_list=syn_qty_all,
        col_name_qty="Ilość_WZ"
    )
else:
    pattern_wz = r"\s*\d+\s+(\d{13})\s+.+?\s+([\d\s]+,\d{2})\s+[\d\s]+,\d{2}$"
    df_wz = parse_pdf_generic(
        uploaded_wz,
        pattern=pattern_wz,
        col_name_qty="Ilość_WZ"
    )

# -------------------------
# 3) Grupowanie i porównanie
# -------------------------
df_ord_g = (
    df_order.groupby("Symbol", as_index=False)
            .agg({"Ilość_Zam":"sum"})
            .rename(columns={"Ilość_Zam":"Zamówiona_ilość"})
)
df_wz_g = (
    df_wz.groupby("Symbol", as_index=False)
         .agg({"Ilość_WZ":"sum"})
         .rename(columns={"Ilość_WZ":"Wydana_ilość"})
)

df_cmp = pd.merge(df_ord_g, df_wz_g, on="Symbol", how="outer", indicator=True)
df_cmp["Zamówiona_ilość"] = df_cmp["Zamówiona_ilość"].fillna(0)
df_cmp["Wydana_ilość"]    = df_cmp["Wydana_ilość"].fillna(0)
df_cmp["Różnica"]         = df_cmp["Zamówiona_ilość"] - df_cmp["Wydana_ilość"]

def status(r):
    if r["_merge"]=="left_only":   return "Brak we WZ"
    if r["_merge"]=="right_only":  return "Brak w zamówieniu"
    return "OK" if r["Różnica"]==0 else "Różni się"

df_cmp["Status"] = df_cmp.apply(status, axis=1)
order = ["Różni się","Brak we WZ","Brak w zamówieniu","OK"]
df_cmp["Status"] = pd.Categorical(df_cmp["Status"], categories=order, ordered=True)
df_cmp.sort_values(["Status","Symbol"], inplace=True)

# -------------------------
# 4) Wyświetlenie i eksport
# -------------------------
def highlight(r):
    color = "#c6efce" if r["Status"]=="OK" else "#ffc7ce"
    return [f"background-color: {color}" for _ in r.index]

st.markdown("### 📊 Wynik porównania")
styled = df_cmp.style.format({
    "Zamówiona_ilość":"{:.0f}",
    "Wydana_ilość":"{:.0f}",
    "Różnica":"{:.0f}"
}).apply(highlight, axis=1)
st.dataframe(styled, use_container_width=True)

def to_excel(df):
    out = BytesIO()
    writer = pd.ExcelWriter(out, engine="openpyxl")
    df.to_excel(writer, index=False, sheet_name="Porównanie")
    writer.close()
    return out.getvalue()

st.download_button("⬇️ Pobierz raport Excel",
    data=to_excel(df_cmp),
    file_name="porownanie_order_vs_wz.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

all_ok = (df_cmp["Status"]=="OK").all()
if all_ok:
    st.markdown("<h4 style='color:green;'>✅ Pozycje się zgadzają</h4>", unsafe_allow_html=True)
else:
    st.markdown("<h4 style='color:red;'>❌ Pozycje się nie zgadzają</h4>", unsafe_allow_html=True)
