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

df_order = pd.DataFrame({
    "Symbol": df_order_raw[col_ean_order].astype(str).str.strip().str.replace(r"\.0+$","",regex=True),
    "Ilość": pd.to_numeric(df_order_raw[col_qty_order], errors="coerce").fillna(0)
})

# 3) Parsowanie WZ
extension = uploaded_wz.name.lower().rsplit(".",1)[-1]

if extension == "pdf":
    try:
        with pdfplumber.open(uploaded_wz) as pdf:
            wz_rows = []

            syn_ean_wz = { normalize_col_name(c): c for c in ["Kod produktu","EAN","symbol"] }
            syn_qty_wz = { normalize_col_name(c): c for c in ["Ilość","Ilosc","Quantity","Qty"] }

            def parse_wz_table(df_table: pd.DataFrame):
                cols = list(df_table.columns)

                # 1) EAN
                col_ean = next((c for c in cols if normalize_col_name(c) in syn_ean_wz), None)
                if not col_ean:
                    return

                # 2) Ilość – prosta kolumna
                col_qty = next((c for c in cols if normalize_col_name(c) in syn_qty_wz), None)
                if col_qty:
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
                    return

                # 3) Broken header: 'Termin ważności Ilość'
                col_part = next((c for c in cols if "termin" in normalize_col_name(c) and "ilo" in normalize_col_name(c)), None)
                if not col_part:
                    return

                for _, row in df_table.iterrows():
                    raw_ean = str(row[col_ean]).strip().split()[-1]
                    if not re.fullmatch(r"\d{13}", raw_ean):
                        continue
                    # Usuwamy separatory tysięcy i zamieniamy przecinek
                    part_cell = str(row[col_part]).strip()
                    part_clean = part_cell.replace(" ", "").replace(",", ".")
                    try:
                        qty = float(part_clean)
                    except:
                        qty = 0.0
                    wz_rows.append([raw_ean, qty])

            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if not table or len(table) < 2:
                        continue
                    hdr = table[0]
                    data = table[1:]
                    df_page = pd.DataFrame(data, columns=hdr)
                    parse_wz_table(df_page)

    except Exception as e:
        st.error(f"Nie udało się przetworzyć PDF:\n```{e}```")
        st.stop()

    if not wz_rows:
        st.error("Nie znaleziono żadnych danych w PDF WZ.")
        st.stop()

    df_wz = pd.DataFrame(wz_rows, columns=["Symbol","Ilość_WZ"])
    df_wz["Symbol"] = df_wz["Symbol"].astype(str).str.strip()
    df_wz["Ilość_WZ"] = pd.to_numeric(df_wz["Ilość_WZ"], errors="coerce").fillna(0)

else:
    # Excelowy WZ
    try:
        df_wz_raw = pd.read_excel(uploaded_wz, dtype=str)
    except Exception as e:
        st.error(f"Nie udało się wczytać Excela WZ:\n```{e}```")
        st.stop()

    syn_ean_wz = { normalize_col_name(c): c for c in ["Kod produktu","EAN","symbol"] }
    syn_qty_wz = { normalize_col_name(c): c for c in ["Ilość","Ilosc","Quantity","Qty"] }

    col_ean_wz = next((c for c in df_wz_raw.columns if normalize_col_name(c) in syn_ean_wz), None)
    col_qty_wz = next((c for c in df_wz_raw.columns if normalize_col_name(c) in syn_qty_wz), None)
    if not col_ean_wz or not col_qty_wz:
        st.error(
            "Excel WZ musi mieć kolumny EAN i Ilość.\n"
            f"Znalezione: {list(df_wz_raw.columns)}"
        )
        st.stop()

    tmp = df_wz_raw[col_ean_wz].astype(str).str.strip().str.split().str[-1]
    mask = tmp.str.fullmatch(r"\d{13}")
    df_wz = pd.DataFrame({
        "Symbol": tmp[mask],
        "Ilość_WZ": pd.to_numeric(
            df_wz_raw.loc[mask, col_qty_wz]
                .astype(str)
                .str.replace(",",".")
                .str.replace(r"\s+","", regex=True),
            errors="coerce"
        ).fillna(0)
    })

# 4) Grupowanie, porównanie i wyświetlenie
