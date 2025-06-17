import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO

st.set_page_config(page_title="Porównywarka Zamówienie vs WZ", layout="wide")
st.title("📋 Porówniwarka Zamówienie (Excel) ↔ WZ (PDF lub Excel)")

def normalize_col_name(name: str) -> str:
    return name.lower().replace(" ", "").replace("\xa0", "").replace("_", "")

def highlight_status_row(row):
    color = "#c6efce" if row["Status"] == "OK" else "#ffc7ce"
    return [f"background-color: {color}" for _ in row.index]

# --- 1) Uploaderzy ---
st.sidebar.header("1. Wgraj pliki")
uploaded_order = st.sidebar.file_uploader("Excel (zamówienie)", type=["xlsx"])
uploaded_wz    = st.sidebar.file_uploader("WZ (PDF lub Excel)", type=["pdf","xlsx"])
if not uploaded_order or not uploaded_wz:
    st.sidebar.info("Potrzebne oba pliki: Excel i PDF/Excel WZ.")
    st.stop()

# --- 2) Parsujemy zamówienie (Excel) ---
df_order_raw = pd.read_excel(uploaded_order, dtype=str)
syn_ean_ord = {normalize_col_name(c): c for c in ["Symbol","symbol","kod ean","ean","kod produktu"]}
syn_qty_ord = {normalize_col_name(c): c for c in ["Ilość","Ilosc","Quantity","Qty","sztuki"]}
def find_col(df, syns):
    for c in df.columns:
        if normalize_col_name(c) in syns:
            return c
    return None

col_ean_o = find_col(df_order_raw, syn_ean_ord)
col_qty_o = find_col(df_order_raw, syn_qty_ord)
if not col_ean_o or not col_qty_o:
    st.error(f"Brak kolumn EAN/Ilość w zamówieniu.\nZnalezione: {list(df_order_raw.columns)}")
    st.stop()

df_order = pd.DataFrame({
    "Symbol": df_order_raw[col_ean_o].astype(str)
                   .str.strip().str.replace(r"\.0+$","",regex=True),
    "Ilość" : pd.to_numeric(df_order_raw[col_qty_o], errors="coerce").fillna(0)
})

# --- 3) Parsujemy WZ ---
extension = uploaded_wz.name.lower().rsplit(".",1)[-1]
wz_rows = []

if extension == "pdf":
    syn_ean_wz = set(normalize_col_name(c) for c in ["Kod produktu","EAN","symbol"])
    syn_qty_wz = set(normalize_col_name(c) for c in ["Ilość","Ilosc","Quantity","Qty"])
    with pdfplumber.open(uploaded_wz) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                if not table or len(table) < 2:
                    continue
                # 3.1) znajdź wiersz nagłówka
                header_idx = None
                norm_header = None
                for i,row in enumerate(table):
                    low = [normalize_col_name(str(c)) for c in row]
                    if any(c in syn_ean_wz for c in low) and any(c in syn_qty_wz for c in low):
                        header_idx = i
                        norm_header = low
                        break
                if header_idx is None:
                    continue
                # indeksy kolumn
                sym_idx = next(j for j,v in enumerate(norm_header) if v in syn_ean_wz)
                qty_idx = next(j for j,v in enumerate(norm_header) if v in syn_qty_wz)
                # 3.2) wyciągnij wiersze danych
                for data_row in table[header_idx+1:]:
                    if len(data_row) <= max(sym_idx, qty_idx):
                        continue
                    raw_ean = str(data_row[sym_idx]).strip().split()[-1]
                    if not re.fullmatch(r"\d{13}", raw_ean):
                        continue
                    qty_cell = str(data_row[qty_idx]).strip()
                    # usuń wszystkie białe znaki, zamień przecinek na kropkę
                    clean_qty = re.sub(r"\s+", "", qty_cell).replace(",", ".")
                    try:
                        qty = float(clean_qty)
                    except:
                        qty = 0.0
                    wz_rows.append([raw_ean, qty])

    if not wz_rows:
        st.error("Nie udało się wyciągnąć danych z PDF WZ.")
        st.stop()

    df_wz = pd.DataFrame(wz_rows, columns=["Symbol","Ilość_WZ"])
    df_wz["Ilość_WZ"] = pd.to_numeric(df_wz["Ilość_WZ"], errors="coerce").fillna(0)

else:
    # Excelowy WZ – analogicznie
    df_wz_raw = pd.read_excel(uploaded_wz, dtype=str)
    syn_ean_wz = {normalize_col_name(c):c for c in ["Kod produktu","EAN","symbol"]}
    syn_qty_wz = {normalize_col_name(c):c for c in ["Ilość","Ilosc","Quantity","Qty"]}
    col_ean_w = next((c for c in df_wz_raw.columns if normalize_col_name(c) in syn_ean_wz), None)
    col_qty_w = next((c for c in df_wz_raw.columns if normalize_col_name(c) in syn_qty_wz), None)
    if not col_ean_w or not col_qty_w:
        st.error(f"Brak kolumn EAN/Ilość w Excelu WZ.\nZnalezione: {list(df_wz_raw.columns)}")
        st.stop()
    tmp = df_wz_raw[col_ean_w].astype(str).str.strip().str.split().str[-1]
    mask = tmp.str.fullmatch(r"\d{13}")
    df_wz = pd.DataFrame({
        "Symbol": tmp[mask],
        "Ilość_WZ": pd.to_numeric(
            df_wz_raw.loc[mask, col_qty_w].astype(str)
                       .apply(lambda x: re.sub(r"\s+","",x).replace(",", ".")),
            errors="coerce"
        ).fillna(0)
    })

# --- 4) Grupowanie i sumowanie ---
df_ord_g = df_order.groupby("Symbol", as_index=False)\
                   .agg(Zamówiona_ilość=("Ilość","sum"))
df_wz_g  = df_wz.groupby("Symbol", as_index=False)\
                .agg(Wydana_ilość=("Ilość_WZ","sum"))

# --- 5) Porównanie ---
df_cmp = pd.merge(df_ord_g, df_wz_g, on="Symbol", how="outer", indicator=True)
df_cmp[["Zamówiona_ilość","Wydana_ilość"]] = df_cmp[["Zamówiona_ilość","Wydana_ilość"]].fillna(0)
df_cmp["Różnica"] = df_cmp["Zamówiona_ilość"] - df_cmp["Wydana_ilość"]

def status(r):
    if r["_merge"]=="left_only":  return "Brak we WZ"
    if r["_merge"]=="right_only": return "Brak w zamówieniu"
    return "OK" if r["Różnica"]==0 else "Różni się"

df_cmp["Status"] = df_cmp.apply(status, axis=1)
order_stats = ["Różni się","Brak we WZ","Brak w zamówieniu","OK"]
df_cmp["Status"] = pd.Categorical(df_cmp["Status"], categories=order_stats, ordered=True)
df_cmp = df_cmp.sort_values(["Status","Symbol"])

# --- 6) Wyświetlenie i eksport ---
st.markdown("### 📊 Wynik porównania")
styled = (
    df_cmp.style
          .format({"Zamówiona_ilość":"{:.0f}","Wydana_ilość":"{:.0f}","Różnica":"{:.0f}"})
          .apply(highlight_status_row, axis=1)
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

if (df_cmp["Status"]=="OK").all():
    st.markdown("<h4 style='color:green;'>✅ Wszystkie pozycje OK</h4>", unsafe_allow_html=True)
else:
    st.markdown("<h4 style='color:red;'>❌ Są różnice!</h4>", unsafe_allow_html=True)
