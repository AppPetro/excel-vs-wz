import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import re

st.set_page_config(
    page_title="📋 Porównywarka Zamówienie ↔ WZ (PDF→Excel)",
    layout="wide",
)

st.title("📋 Porównywarka Zamówienie (Excel) vs. WZ (PDF lub Excel)")

st.markdown(
    """
    **Instrukcja:**
    1. Wgraj Excel z zamówieniem, zawierający kolumny z nazwami EAN i ilości:
       - EAN: `Symbol`, `symbol`, `kod ean`, `ean`, `kod produktu`
       - Ilość: `Ilość`, `Ilosc`, `Quantity`, `Qty`, `sztuki`
    2. Wgraj WZ w formie **PDF** (lub Excel), gdzie kolumna EAN może się nazywać:
       - `Kod produktu`, `EAN`, `symbol`
       - Ilość: `Ilość`, `Ilosc`, `Quantity`, `Qty`
    3. Aplikacja:
       - rozpozna synonimy kolumn,
       - z PDF → przeprocesuje `extract_tables()`,
       - zsumuje po EAN-ach i porówna z zamówieniem,
       - wyświetli tabelę z kolorowaniem i pozwoli pobrać wynik.
    """
)

def highlight_status_row(row):
    color = "#c6efce" if row["Status"] == "OK" else "#ffc7ce"
    return [f"background-color: {color}" for _ in row.index]

def normalize_col_name(name: str) -> str:
    return name.lower().replace(" ", "").replace("\xa0", "").replace("_", "")

# 1) Wgrywanie plików
st.sidebar.header("Krok 1: Excel (zamówienie)")
uploaded_order = st.sidebar.file_uploader("Wybierz plik zamówienia", type=["xlsx"])
st.sidebar.header("Krok 2: WZ (PDF lub Excel)")
uploaded_wz = st.sidebar.file_uploader("Wybierz plik WZ", type=["pdf", "xlsx"])

if not uploaded_order or not uploaded_wz:
    st.info("Proszę wgrać oba pliki: Excel (zamówienie) oraz PDF/Excel (WZ).")
    st.stop()

# 2) Parsowanie zamówienia
try:
    df_order_raw = pd.read_excel(uploaded_order, dtype=str)
except Exception as e:
    st.error(f"Nie udało się wczytać pliku zamówienia:\n```{e}```")
    st.stop()

syn_ean_ord = { normalize_col_name(c): c for c in ["Symbol","symbol","kod ean","ean","kod produktu"] }
syn_qty_ord = { normalize_col_name(c): c for c in ["Ilość","Ilosc","Quantity","Qty","sztuki"] }

def find_col(df, syns):
    for c in df.columns:
        if normalize_col_name(c) in syns:
            return c
    return None

col_ean_order = find_col(df_order_raw, syn_ean_ord)
col_qty_order = find_col(df_order_raw, syn_qty_ord)
if not col_ean_order or not col_qty_order:
    st.error(
        "Excel zamówienia musi mieć kolumny EAN i Ilość.\n"
        f"Znalezione: {list(df_order_raw.columns)}"
    )
    st.stop()

# przygotowanie DataFrame zamówienia
df_order = pd.DataFrame({
    "Symbol": df_order_raw[col_ean_order].astype(str).str.strip().str.replace(r"\.0+$","",regex=True),
    "Ilość": pd.to_numeric(df_order_raw[col_qty_order], errors="coerce").fillna(0)
})

# 3) Parsowanie WZ
ext = uploaded_wz.name.lower().split('.')[-1]
if ext == 'pdf':
    with pdfplumber.open(uploaded_wz) as pdf:
        wz_rows = []
        syn_ean_wz = { normalize_col_name(c): c for c in ["Kod produktu","EAN","symbol"] }
        syn_qty_wz = { normalize_col_name(c): c for c in ["Ilość","Ilosc","Quantity","Qty"] }
        
        def parse_wz_table(df_table: pd.DataFrame):
            cols = list(df_table.columns)
            col_ean = next((c for c in cols if normalize_col_name(c) in syn_ean_wz), None)
            col_qty = next((c for c in cols if normalize_col_name(c) in syn_qty_wz), None)
            if not col_ean or not col_qty:
                return
            for _, row in df_table.iterrows():
                raw_ean = str(row[col_ean]).strip().split()[-1]
                if not re.fullmatch(r"\d{13}", raw_ean):
                    continue
                raw_qty = str(row[col_qty]).replace(" ","").replace(",",".")
                try:
                    qty = float(raw_qty)
                except:
                    qty = 0.0
                wz_rows.append([raw_ean, qty])

        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if table and len(table) > 1:
                    df_page = pd.DataFrame(table[1:], columns=table[0])
                    parse_wz_table(df_page)

    df_wz = pd.DataFrame(wz_rows, columns=["Symbol","Ilość_WZ"]).groupby("Symbol", as_index=False).sum()
else:
    df_wz_raw = pd.read_excel(uploaded_wz, dtype=str)
    syn_ean_wz = { normalize_col_name(c): c for c in ["Kod produktu","EAN","symbol"] }
    syn_qty_wz = { normalize_col_name(c): c for c in ["Ilość","Ilosc","Quantity","Qty"] }
    col_ean_wz = find_col(df_wz_raw, syn_ean_wz)
    col_qty_wz = find_col(df_wz_raw, syn_qty_wz)
    df_wz = pd.DataFrame({
        "Symbol": df_wz_raw[col_ean_wz].astype(str).str.strip().str.split().str[-1],
        "Ilość_WZ": pd.to_numeric(
            df_wz_raw[col_qty_wz].astype(str).str.replace(",",".").str.replace(r"\s+","",regex=True),
            errors="coerce"
        ).fillna(0)
    }).groupby("Symbol", as_index=False).sum()

# 4) Porównanie
df_cmp = pd.merge(
    df_order.groupby("Symbol", as_index=False).agg({"Ilość":"sum"}).rename(columns={"Ilość":"Zamówiona"}),
    df_wz.rename(columns={"Ilość_WZ":"Wydana"}),
    on="Symbol", how="outer", indicator=True
)
df_cmp["Zamówiona"] = df_cmp["Zamówiona"].fillna(0)
df_cmp["Wydana"] = df_cmp["Wydana"].fillna(0)
df_cmp["Różnica"] = df_cmp["Zamówiona"] - df_cmp["Wydana"]
def stat(r):
    if r["_merge"]=="left_only": return "Brak we WZ"
    if r["_merge"]=="right_only": return "Brak w zamówieniu"
    return "OK" if r["Różnica"]==0 else "Różni się"
df_cmp["Status"] = df_cmp.apply(stat, axis=1)
order_stats = ["Różni się","Brak we WZ","Brak w zamówieniu","OK"]
df_cmp["Status"] = pd.Categorical(df_cmp["Status"], categories=order_stats, ordered=True)
df_cmp.sort_values(["Status","Symbol"], inplace=True)

# 5) Wyświetlenie i eksport
st.markdown("### 📊 Wyniki porównania")
st.dataframe(
    df_cmp.style.format({"Zamówiona":"{:.0f}","Wydana":"{:.0f}","Różnica":"{:.0f}"}).apply(highlight_status_row, axis=1),
    use_container_width=True
)

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

# Podsumowanie
if (df_cmp["Status"]=="OK").all():
    st.markdown("<h4 style='color:green;'>✅ Pozycje się zgadzają</h4>", unsafe_allow_html=True)
else:
    st.markdown("<h4 style='color:red;'>❌ Pozycje się nie zgadzają</h4>", unsafe_allow_html=True)
