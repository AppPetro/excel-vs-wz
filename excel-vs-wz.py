import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import re

st.set_page_config(
    page_title="📋 Porównywarka Zamówienie/Zlecenie ↔ WZ (PDF/Excel)",
    layout="wide",
)

st.title("📋 Porównywarka Zamówienie/Zlecenie (PDF lub Excel) vs. WZ (PDF lub Excel)")
st.markdown(
    """
    **Instrukcja:**
    1. Wgraj Zlecenie (Excel lub PDF), zawierające pozycje z kolumnami EAN i ilości.
    2. Wgraj WZ (PDF lub Excel), gdzie kolumna EAN może się nazywać:
       - `Kod produktu`, `EAN`, `symbol`
       - Ilość: `Ilość`, `Ilosc`, `Quantity`, `Qty`
    3. Aplikacja:
       - rozpozna synonimy kolumn lub użyje regex do PDF,
       - zsumuje po EAN-ach i porówna z WZ,
       - wyświetli tabelę z kolorowaniem i pozwoli pobrać wynik.
    """
)

def highlight_status_row(row):
    color = "#c6efce" if row["Status"] == "OK" else "#ffc7ce"
    return [f"background-color: {color}" for _ in row.index]

def normalize_col_name(name: str) -> str:
    return name.lower().replace(" ", "").replace("\xa0", "").replace("_", "")

# -------------------------
# 1) Wgrywanie plików
# -------------------------
st.sidebar.header("Krok 1: Zlecenie/Zamówienie (PDF lub Excel)")
uploaded_order = st.sidebar.file_uploader("Wybierz Zlecenie/Zamówienie", type=["xlsx","pdf"])
st.sidebar.header("Krok 2: WZ (PDF lub Excel)")
uploaded_wz = st.sidebar.file_uploader("Wybierz WZ", type=["pdf", "xlsx"])

if not uploaded_order or not uploaded_wz:
    st.info("Proszę wgrać oba pliki: Zlecenie/Zamówienie oraz PDF/Excel (WZ).")
    st.stop()

# -------------------------
# 2) Parsowanie Zlecenia/Zamówienia
# -------------------------
ext_ord = uploaded_order.name.lower().rsplit(".",1)[-1]
if ext_ord == "pdf":
    try:
        ord_rows = []
        with pdfplumber.open(uploaded_order) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                for line in text.split("\n"):
                    # numer pozycji + EAN(13) + opis + ilość
                    m = re.match(
                        r"\s*\d+\s+(\d{13})\s+.+?\s+([\d\s]+,\d{2})$",
                        line
                    )
                    if not m:
                        continue
                    raw_ean = m.group(1)
                    raw_qty = m.group(2).replace(" ", "").replace(",", ".")
                    try:
                        qty = float(raw_qty)
                    except:
                        qty = 0.0
                    ord_rows.append([raw_ean, qty])
        if not ord_rows:
            st.error("Nie znaleziono danych w PDF Zlecenia/Zamówienia.")
            st.stop()
        df_order = pd.DataFrame(ord_rows, columns=["Symbol","Ilość"])
    except Exception as e:
        st.error(f"Nie udało się przetworzyć PDF Zlecenia/Zamówienia:\n```{e}```")
        st.stop()
else:
    try:
        df_order_raw = pd.read_excel(uploaded_order, dtype=str)
    except Exception as e:
        st.error(f"Nie udało się wczytać pliku Zlecenia/Zamówienia:\n```{e}```")
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
            "Excel Zlecenia/Zamówienia musi mieć kolumny EAN i Ilość."
            f" Znalezione: {list(df_order_raw.columns)}"
        )
        st.stop()
    df_order = pd.DataFrame({
        "Symbol": df_order_raw[col_ean_order].astype(str).str.strip().str.replace(r"\.0+$","",regex=True),
        "Ilość": pd.to_numeric(df_order_raw[col_qty_order], errors="coerce").fillna(0)
    })

# -------------------------
# 3) Parsowanie WZ (jak poprzednio)
# -------------------------
ext_wz = uploaded_wz.name.lower().rsplit(".",1)[-1]
if ext_wz == "pdf":
    try:
        wz_rows = []
        with pdfplumber.open(uploaded_wz) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                for line in text.split("\n"):
                    m = re.match(
                        r"\s*\d+\s+(\d{13})\s+.+?\s+([\d\s]+,\d{2})\s+[\d\s]+,\d{2}$",
                        line
                    )
                    if not m:
                        continue
                    raw_ean = m.group(1)
                    raw_qty = m.group(2).replace(" ", "").replace(",", ".")
                    try:
                        qty = float(raw_qty)
                    except:
                        qty = 0.0
                    wz_rows.append([raw_ean, qty])
        if not wz_rows:
            st.error("Nie znaleziono danych w PDF WZ.")
            st.stop()
        df_wz = pd.DataFrame(wz_rows, columns=["Symbol","Ilość_WZ"]).
            astype({"Symbol":str}).assign(
                Ilość_WZ=lambda df:df["Ilość_WZ"]
            )
    except Exception as e:
        st.error(f"Nie udało się przetworzyć PDF WZ:\n```{e}```")
        st.stop()
else:
    # kod Excelowego WZ bez zmian
    ...
# resto pozostaje tak jak wcześniej (grupowanie, porównanie, wyświetlenie)
