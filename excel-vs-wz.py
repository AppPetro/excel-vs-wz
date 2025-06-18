import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import re

st.set_page_config(
    page_title="📋 Porównywarka Zamówienie ↔ WZ",
    layout="wide",
)

st.title("📋 Porównywarka Zamówienie (Excel) vs. WZ (PDF/Excel)")

# ------------------------------
# Funkcje pomocnicze
# ------------------------------
def normalize_col(name: str) -> str:
    # usuń spacje, twarde spacje, podkreślniki, kropki i przemnóż na lowercase
    return re.sub(r"[\s\xa0_\.]+", "", str(name)).lower()

def highlight_status(row):
    color = "#c6efce" if row.Status == "OK" else "#ffc7ce"
    return [f"background-color: {color}"] * len(row)

# ------------------------------
# 1) Wczytanie plików
# ------------------------------
st.sidebar.header("1️⃣ Wczytaj pliki")
file_order = st.sidebar.file_uploader("Excel zamówienia", type=["xlsx"])
file_wz    = st.sidebar.file_uploader("WZ (PDF lub Excel)", type=["pdf","xlsx"])
if not file_order or not file_wz:
    st.info("Proszę wgrać oba pliki: Zamówienie ➡️ WZ po lewej stronie.")
    st.stop()

# ------------------------------
# 2) Parsowanie Excela zamówienia
# ------------------------------
df_order_raw = pd.read_excel(file_order, dtype=str)
# synonimy kolumn
syn_ean_ord = {normalize_col(c): c for c in ["Symbol","symbol","kod ean","ean","kod produktu"]}
syn_qty_ord = {normalize_col(c): c for c in ["ilość","ilosc","quantity","qty","sztuki"]}
# funkcja wyszukująca
 def find_col(df, syn):
    for c in df.columns:
        if normalize_col(c) in syn:
            return c
    return None
# znajdź kolumny
col_ean_ord = find_col(df_order_raw, syn_ean_ord)
col_qty_ord = find_col(df_order_raw, syn_qty_ord)
if not col_ean_ord or not col_qty_ord:
    st.error(f"Brak kolumn EAN lub Ilość w pliku zamówienia. Znalezione: {list(df_order_raw.columns)}")
    st.stop()
# budowa df_order
df_order = pd.DataFrame({
    "Symbol": df_order_raw[col_ean_ord].astype(str).str.extract(r"(\d{13})")[0],
    "Zamówiona": df_order_raw[col_qty_ord].astype(str)
        .str.replace(r"[\s\.]+", "", regex=True)
        .str.replace(",", ".")
})
df_order["Zamówiona"] = pd.to_numeric(df_order["Zamówiona"], errors="coerce").fillna(0)
df_order = df_order.groupby("Symbol", as_index=False).sum()

# ------------------------------
# 3) Parsowanie WZ
# ------------------------------
syn_ean_wz = syn_ean_ord.copy()
syn_qty_wz = syn_qty_ord.copy()
ext = file_wz.name.lower().rsplit('.',1)[-1]
rows = []
if ext == "pdf":
    with pdfplumber.open(file_wz) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables() or []
            for table in tables:
                if len(table) < 2:
                    continue
                header = table[0]
                # lokalizuj indeksy kolumn EAN i Ilość
                ean_idx = next((i for i, c in enumerate(header) if normalize_col(c) in syn_ean_wz), None)
                qty_idx = next((i for i, c in enumerate(header) if normalize_col(c) in syn_qty_wz), None)
                if ean_idx is None or qty_idx is None:
                    continue  # pomijaj tabele bez wymaganych kolumn
                data = table[1:]
                for row in data:
                    raw_ean = str(row[ean_idx]).strip()
                    m = re.search(r"(\d{13})", raw_ean)
                    if not m:
                        continue
                    ean = m.group(1)
                    raw_qty = str(row[qty_idx]).strip()
                    # oczyszczanie
                    clean = raw_qty.replace(" ","").replace("\xa0","").replace(",",".")
                    try:
                        qty = float(clean)
                    except:
                        qty = 0.0
                    rows.append((ean, qty))
else:
    df_wz_raw = pd.read_excel(file_wz, dtype=str)
    # znajdź kolumny EAN i ilość
    col_ean_wz = find_col(df_wz_raw, syn_ean_wz)
    col_qty_wz = find_col(df_wz_raw, syn_qty_wz)
    if not col_ean_wz or not col_qty_wz:
        st.error(f"Brak kolumn EAN lub Ilość w pliku WZ. Znalezione: {list(df_wz_raw.columns)}")
        st.stop()
    for _, r in df_wz_raw.iterrows():
        raw_ean = str(r[col_ean_wz]).strip()
        m = re.search(r"(\d{13})", raw_ean)
        if not m:
            continue
        ean = m.group(1)
        raw_qty = str(r[col_qty_wz]).strip()
        clean = raw_qty.replace(" ","").replace("\xa0","").replace(",",".")
        try:
            qty = float(clean)
        except:
            qty = 0.0
        rows.append((ean, qty))
# tworzenie df_wz
if not rows:
    st.error("Nie znaleziono żadnych danych w WZ.")
    st.stop()
df_wz = pd.DataFrame(rows, columns=["Symbol","Wydana"]).groupby("Symbol", as_index=False).sum()

# ------------------------------
# 4) Porównanie
# ------------------------------
df_cmp = pd.merge(df_order, df_wz, on="Symbol", how="outer", indicator=True)
df_cmp[["Zamówiona","Wydana"]] = df_cmp[["Zamówiona","Wydana"]].fillna(0)
df_cmp["Różnica"] = df_cmp["Zamówiona"] - df_cmp["Wydana"]
def get_status(row):
    if row._merge == 'left_only': return 'Brak we WZ'
    if row._merge == 'right_only': return 'Brak w zamówieniu'
    return 'OK' if row.Różnica == 0 else 'Różni się'
df_cmp['Status'] = df_cmp.apply(get_status, axis=1)
order_cat = ['Różni się','Brak we WZ','Brak w zamówieniu','OK']
df_cmp['Status'] = pd.Categorical(df_cmp['Status'], categories=order_cat, ordered=True)
df_cmp.sort_values(['Status','Symbol'], inplace=True)

# ------------------------------
# 5) Wyświetlenie i eksport
# ------------------------------
st.markdown("### 📊 Wyniki porównania")
st.dataframe(
    df_cmp.style
        .format({'Zamówiona':'{:.0f}','Wydana':'{:.0f}','Różnica':'{:.0f}'})
        .apply(highlight_status, axis=1),
    use_container_width=True
)

def to_excel(df):
    buf = BytesIO()
    writer = pd.ExcelWriter(buf, engine='openpyxl')
    df.to_excel(writer, index=False, sheet_name='Porównanie')
    writer.close()
    return buf.getvalue()

st.download_button(
    label='⬇️ Pobierz raport',
    data=to_excel(df_cmp),
    file_name='porownanie_order_vs_wz.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)

# ------------------------------
# Podsumowanie
# ------------------------------
if (df_cmp['Status'] == 'OK').all():
    st.success("✅ Wszystkie pozycje się zgadzają")
else:
    st.error("❌ Wykryto rozbieżności w pozycjach")
