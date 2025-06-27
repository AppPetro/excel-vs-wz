import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import re

st.set_page_config(
    page_title="📋 Porównywarka Zlecenie/Zamówienie ↔ WZ (PDF/Excel)",
    layout="wide",
)

st.title("📋 Porównywarka Zlecenie/Zamówienie (PDF/Excel) vs. WZ (PDF/Excel)")
st.markdown(
    """
    **Instrukcja:**
    1. Wgraj Zlecenie/Zamówienie (PDF lub Excel), zawierające kolumny EAN i ilości.
    2. Wgraj WZ (PDF lub Excel), gdzie kolumna EAN może się nazywać:
       - `Kod produktu`, `EAN`, `symbol`
       - Ilość: `Ilość`, `Ilosc`, `Quantity`, `Qty`
    3. Program przetworzy oba pliki (PDF regex lub Excel kolumny), zsumuje ilości po EAN-ach i porówna.
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
uploaded_order = st.sidebar.file_uploader(
    "Wybierz plik Zlecenia/Zamówienia", type=["xlsx", "pdf"]
)
st.sidebar.header("Krok 2: WZ (PDF lub Excel)")
uploaded_wz = st.sidebar.file_uploader(
    "Wybierz plik WZ", type=["xlsx", "pdf"]
)

if not uploaded_order or not uploaded_wz:
    st.info("Proszę wgrać oba pliki: Zlecenie/Zamówienie oraz WZ.")
    st.stop()


# -------------------------
# 2) Parsowanie Zlecenia/Zamówienia
# -------------------------
ext_order = uploaded_order.name.lower().rsplit(".", 1)[-1]
if ext_order == "pdf":
    # PDF: regex scan ilość -> jednostka -> EAN
    ord_rows = []
    try:
        with pdfplumber.open(uploaded_order) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                for line in text.split("\n"):
                    m = re.match(r"\s*\d+\s+.*?\s+([\d\s]+,\d+)\s+\S+\s+(\d{13})", line)
                    if not m:
                        continue
                    raw_q, raw_ean = m.group(1), m.group(2)
                    raw_qty = re.sub(r"\s+", "", raw_q).replace(",", ".")
                    try:
                        qty = float(raw_qty)
                    except:
                        qty = 0.0
                    ord_rows.append([raw_ean, qty])
        if not ord_rows:
            st.error("Nie znaleziono danych w PDF Zlecenia/Zamówienia.")
            st.stop()
        df_order = pd.DataFrame(ord_rows, columns=["Symbol", "Ilość"]).astype({"Symbol": str, "Ilość": float})
    except Exception as e:
        st.error(f"Błąd przetwarzania PDF Zlecenia/Zamówienia:\n```{e}```")
        st.stop()
else:
    # Excel: znajdź wiersz nagłówka gdzie są EAN+iLość (może być w 15 wierszu)
    try:
        df_temp = pd.read_excel(uploaded_order, dtype=str, header=None)
    except Exception as e:
        st.error(f"Błąd wczytywania Excela Zlecenia/Zamówienia:\n```{e}```")
        st.stop()
    # Synonimy kolumn
    syn_ean_ord = {normalize_col_name(c): c for c in [
        "EAN", "Kod EAN", "GTIN", "kod ean", "kod produktu", "symbol"
    ]}
    syn_qty_ord = {normalize_col_name(c): c for c in [
        "Ilość", "Ilosc", "Quantity", "Qty", "sztuki", "ilość sztuk", "ilość sztuk zamówiona", "zamówiona ilość", "quantity ordered", "qty ordered"
    ]}
    header_idx = None
    for idx, row in df_temp.iterrows():
        cols = [normalize_col_name(str(x)) for x in row.tolist()]
        if any(c in syn_ean_ord for c in cols) and any(c in syn_qty_ord for c in cols):
            header_idx = idx
            break
    if header_idx is None:
        st.error("Excel Zlecenia/Zamówienia musi mieć w nagłówku kolumny EAN i Ilość.")
        st.stop()
    df_order_raw = df_temp.iloc[header_idx+1:].copy().reset_index(drop=True)
    df_order_raw.columns = df_temp.iloc[header_idx]
    # znajdź kolumny i zbuduj df_order
    def find_col(df, syns):
        for c in df.columns:
            if normalize_col_name(c) in syns:
                return c
        return None
    col_ean_order = find_col(df_order_raw, syn_ean_ord)
    col_qty_order = find_col(df_order_raw, syn_qty_ord)
    if not col_ean_order or not col_qty_order:
        st.error(f"Nie znaleziono kolumn EAN/Ilość w Excelu: {list(df_order_raw.columns)}")
        st.stop()
    df_order = pd.DataFrame({
        "Symbol": df_order_raw[col_ean_order].astype(str).str.strip().str.replace(r"\.0+$", "", regex=True),
        "Ilość": pd.to_numeric(
            df_order_raw[col_qty_order].astype(str).str.replace(r"\s+", "", regex=True).str.replace(",", "."),
            errors="coerce"
        ).fillna(0)
    })

# -------------------------
# 3) Parsowanie WZ
# -------------------------
ext_wz = uploaded_wz.name.lower().rsplit(".", 1)[-1]
if ext_wz == "pdf":
    wz_rows = []
    try:
        with pdfplumber.open(uploaded_wz) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                for line in text.split("\n"):
                    m = re.match(r"\s*\d+\s+(\d{13})\s+.*?\s+([\d\s]+,\d{2})\s+[\d\s]+,\d{2}$", line)
                    if not m:
                        continue
                    raw_ean, raw_q = m.group(1), m.group(2)
                    raw_qty = re.sub(r"\s+", "", raw_q).replace(",", ".")
                    try:
                        qty = float(raw_qty)
                    except:
                        qty = 0.0
                    wz_rows.append([raw_ean, qty])
        if not wz_rows:
            st.error("Nie znaleziono danych w PDF WZ.")
            st.stop()
        df_wz = pd.DataFrame(wz_rows, columns=["Symbol", "Ilość_WZ"]).astype({"Symbol": str, "Ilość_WZ": float})
    except Exception as e:
        st.error(f"Błąd przetwarzania PDF WZ:\n```{e}```")
        st.stop()
else:
    try:
        df_wz_raw = pd.read_excel(uploaded_wz, dtype=str)
    except Exception as e:
        st.error(f"Błąd wczytywania Excela WZ:\n```{e}```")
        st.stop()
    syn_ean_wz = {normalize_col_name(c): c for c in ["Kod produktu","EAN","symbol"]}
    syn_qty_wz = {normalize_col_name(c): c for c in ["Ilość","Ilosc","Quantity","Qty"]}
    def find_col(df, syns):
        for c in df.columns:
            if normalize_col_name(c) in syns:
                return c
        return None
    col_ean_wz = find_col(df_wz_raw, syn_ean_wz)
    col_qty_wz = find_col(df_wz_raw, syn_qty_wz)
    if not col_ean_wz or not col_qty_wz:
        st.error(f"Excel WZ musi mieć kolumny EAN i Ilość. Znalezione: {list(df_wz_raw.columns)}")
        st.stop()
    df_wz = pd.DataFrame({
        "Symbol": df_wz_raw[col_ean_wz].astype(str).str.strip().str.split().str[-1],
        "Ilość_WZ": pd.to_numeric(df_wz_raw[col_qty_wz].astype(str).str.replace(r"\s+", "", regex=True).
            str.replace(",", "."), errors="coerce").fillna(0)
    })

# -------------------------
# 4) Grupowanie, sumowanie, porównanie i wyświetlenie
# -------------------------
df_ord_g = df_order.groupby("Symbol", as_index=False).agg({"Ilość": "sum"}).rename(columns={"Ilość": "Zamówiona_ilość"})
df_wz_g  = df_wz.groupby("Symbol", as_index=False).agg({"Ilość_WZ": "sum"}).rename(columns={"Ilość_WZ": "Wydana_ilość"})

# Porównanie
df_cmp = pd.merge(df_ord_g, df_wz_g, on="Symbol", how="outer", indicator=True)
df_cmp["Zamówiona_ilość"] = df_cmp["Zamówiona_ilość"].fillna(0)
df_cmp["Wydana_ilość"]    = df_cmp["Wydana_ilość"].fillna(0)
df_cmp["Różnica"]         = df_cmp["Zamówiona_ilość"] - df_cmp["Wydana_ilość"]

# Status
def status(r):
    if r["_merge"] == "left_only":
        return "Brak we WZ"
    if r["_merge"] == "right_only":
        return "Brak w zamówieniu"
    return "OK" if r["Różnica"] == 0 else "Różni się"

order_stats = ["Różni się", "Brak we WZ", "Brak w zamówieniu", "OK"]
df_cmp["Status"] = pd.Categorical(df_cmp.apply(status, axis=1), categories=order_stats, ordered=True)
df_cmp = df_cmp.sort_values(["Status", "Symbol"])

# Wyświetlenie
st.markdown("### 📊 Wynik porównania")
st.dataframe(
    df_cmp.style
        .format({"Zamówiona_ilość": "{:.0f}", "Wydana_ilość": "{:.0f}", "Różnica": "{:.0f}"})
        .apply(highlight_status_row, axis=1),
    use_container_width=True
)

# Eksport i komunikat zgodności
out = BytesIO()
writer = pd.ExcelWriter(out, engine="openpyxl")
df_cmp.to_excel(writer, index=False, sheet_name="Porównanie")
writer.close()
st.download_button(
    "⬇️ Pobierz raport Excel",
    data=out.getvalue(),
    file_name="porownanie.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
if df_cmp["Status"].eq("OK").all():
    st.markdown("<h4 style='color:green;'>✅ Pozycje się zgadzają</h4>", unsafe_allow_html=True)
else:
    st.markdown("<h4 style='color:red;'>❌ Pozycje się nie zgadzają</h4>", unsafe_allow_html=True)
