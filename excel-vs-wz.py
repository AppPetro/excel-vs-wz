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

# Funkcje pomocnicze
def highlight_status_row(row):
    color = "#c6efce" if row["Status"] == "OK" else "#ffc7ce"
    return [f"background-color: {color}" for _ in row.index]

def normalize_col_name(name: str) -> str:
    return name.lower().replace(" ", "").replace("\xa0", "").replace("_", "")

def find_col(df, syns):
    for c in df.columns:
        if normalize_col_name(c) in syns:
            return c
    return None

# 1) Wczytanie plików
st.sidebar.header("Krok 1: Excel (zamówienie)")
uploaded_order = st.sidebar.file_uploader("Wybierz plik zamówienia", type=["xlsx"])
st.sidebar.header("Krok 2: WZ (PDF lub Excel)")
uploaded_wz = st.sidebar.file_uploader("Wybierz plik WZ", type=["pdf", "xlsx"])
if not uploaded_order or not uploaded_wz:
    st.info("Proszę wgrać oba pliki po lewej.")
    st.stop()

# 2) Parsowanie zamówienia
df_order_raw = pd.read_excel(uploaded_order, dtype=str)
syn_ean_ord = { normalize_col_name(c): c for c in ["Symbol","symbol","kod ean","ean","kod produktu"] }
syn_qty_ord = { normalize_col_name(c): c for c in ["Ilość","Ilosc","Quantity","Qty","sztuki"] }
col_ean_order = find_col(df_order_raw, syn_ean_ord)
col_qty_order = find_col(df_order_raw, syn_qty_ord)
if not col_ean_order or not col_qty_order:
    st.error(f"Brak kolumn EAN/Ilość w zamówieniu: {list(df_order_raw.columns)}")
    st.stop()

df_order = pd.DataFrame({
    "Symbol": df_order_raw[col_ean_order].astype(str).str.strip().str.replace(r"\.0+$", "", regex=True),
    "Ilość": pd.to_numeric(df_order_raw[col_qty_order], errors="coerce").fillna(0)
})

# 3) Parsowanie WZ
ext = uploaded_wz.name.lower().rsplit(".", 1)[-1]
wz_rows = []
syn_ean_wz = { normalize_col_name(c): c for c in ["Kod produktu","EAN","symbol"] }
syn_qty_wz = { normalize_col_name(c): c for c in ["Ilość","Ilosc","Quantity","Qty"] }

if ext == 'pdf':
    with pdfplumber.open(uploaded_wz) as pdf:
        def parse_wz_table(df_table: pd.DataFrame):
            cols = list(df_table.columns)
            # EAN
            col_ean = next((c for c in cols if normalize_col_name(c) in syn_ean_wz), None)
            if not col_ean:
                return
            # Ilość
            col_qty = next((c for c in cols if normalize_col_name(c) in syn_qty_wz), None)
            if not col_qty:
                for rc in cols:
                    low = normalize_col_name(rc)
                    if "termin" in low and "ilo" in low:
                        col_qty = rc
                        break
            if not col_qty:
                return
            # Parsowanie wierszy
            for _, row in df_table.iterrows():
                raw_ean = str(row[col_ean]).strip().split()[-1]
                if not re.fullmatch(r"\d{13}", raw_ean):
                    continue
                raw_qty = str(row[col_qty]).strip().replace(" ", "").replace(",", ".")
                try:
                    qty = float(raw_qty)
                except:
                    qty = 0.0
                wz_rows.append([raw_ean, qty])

        for page in pdf.pages:
            tables = page.extract_tables() or []
            for table in tables:
                if not table or len(table) < 2:
                    continue
                hdr0, hdr1 = table[0], table[1]
                norm0 = [normalize_col_name(str(x)) for x in hdr0]
                norm1 = [normalize_col_name(str(x)) for x in hdr1]
                has_ean0 = any(k in syn_ean_wz for k in norm0)
                has_qty0 = any(k in syn_qty_wz for k in norm0)
                has_ean1 = any(k in syn_ean_wz for k in norm1)
                has_qty1 = any(k in syn_qty_wz for k in norm1)
                if has_ean0 and has_qty0:
                    header, data = hdr0, table[1:]
                elif has_ean1 and has_qty1:
                    header, data = hdr1, table[2:]
                elif has_ean0:
                    header, data = hdr0, table[1:]
                elif has_ean1:
                    header, data = hdr1, table[2:]
                else:
                    continue
                df_page = pd.DataFrame(data, columns=header)
                parse_wz_table(df_page)
else:
    df_wz_raw = pd.read_excel(uploaded_wz, dtype=str)
    col_ean_wz = find_col(df_wz_raw, syn_ean_wz)
    col_qty_wz = find_col(df_wz_raw, syn_qty_wz)
    if not col_ean_wz or not col_qty_wz:
        st.error(f"Brak kolumn EAN/Ilość w pliku WZ: {list(df_wz_raw.columns)}")
        st.stop()
    tmp = df_wz_raw[col_ean_wz].astype(str).str.strip().str.split().str[-1]
    mask = tmp.str.fullmatch(r"\d{13}")
    df_wz = pd.DataFrame({
        "Symbol": tmp[mask],
        "Ilość_WZ": pd.to_numeric(
            df_wz_raw.loc[mask, col_qty_wz].astype(str)
                        .str.replace(",", ".")
                        .str.replace(r"\s+", "", regex=True),
            errors="coerce"
        ).fillna(0)
    })

# 4) Porównanie
if ext == 'pdf':
    df_wz = pd.DataFrame(wz_rows, columns=["Symbol","Ilość_WZ"]).groupby("Symbol", as_index=False).sum()

df_cmp = pd.merge(
    df_order.groupby("Symbol", as_index=False).agg({"Ilość":"sum"}).rename(columns={"Ilość":"Zamówiona_ilość"}),
    df_wz.rename(columns={"Ilość_WZ":"Wydana_ilość"}),
    on="Symbol", how="outer", indicator=True
)
df_cmp["Zamówiona_ilość"].fillna(0, inplace=True)
df_cmp["Wydana_ilość"].fillna(0, inplace=True)
df_cmp["Różnica"] = df_cmp["Zamówiona_ilość"] - df_cmp["Wydana_ilość"]

def status(r):
    if r["_merge"] == "left_only": return "Brak we WZ"
    if r["_merge"] == "right_only": return "Brak w zamówieniu"
    return "OK" if r["Różnica"] == 0 else "Różni się"

df_cmp["Status"] = df_cmp.apply(status, axis=1)
order_stats = ["Różni się","Brak we WZ","Brak w zamówieniu","OK"]
df_cmp["Status"] = pd.Categorical(df_cmp["Status"], categories=order_stats, ordered=True)
df_cmp.sort_values(["Status","Symbol"], inplace=True)

# 5) Wyświetlenie i eksport
st.markdown("### 📊 Wynik porównania")
st.dataframe(
    df_cmp.style.format({"Zamówiona_ilość":"{:.0f}","Wydana_ilość":"{:.0f}","Różnica":"{:.0f}"}).apply(highlight_status_row, axis=1),
    use_container_width=True
)

def to_excel(df):
    buf = BytesIO()
    writer = pd.ExcelWriter(buf, engine="openpyxl")
    df.to_excel(writer, index=False, sheet_name="Porównanie")
    writer.close()
    return buf.getvalue()

st.download_button("⬇️ Pobierz raport", data=to_excel(df_cmp), file_name="porownanie_order_vs_wz.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Podsumowanie
if (df_cmp["Status"] == "OK").all():
    st.success("✅ Wszystkie pozycje się zgadzają")
else:
    st.error("❌ Wykryto rozbieżności")
