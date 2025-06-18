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
    return re.sub(r"[\s\xa0_\.]+", "", str(name)).lower()

def highlight_status(row):
    color = "#c6efce" if row.Status == "OK" else "#ffc7ce"
    return [f"background-color: {color}"] * len(row)

def find_col(df: pd.DataFrame, synonyms: set) -> str:
    for col in df.columns:
        if normalize_col(col) in synonyms:
            return col
    return None

# ------------------------------
# 1) Wczytanie plików
# ------------------------------
st.sidebar.header("1️⃣ Wczytaj pliki")
file_order = st.sidebar.file_uploader("Excel zamówienia", type=["xlsx"])
file_wz    = st.sidebar.file_uploader("WZ (PDF lub Excel)", type=["pdf","xlsx"])
if not file_order or not file_wz:
    st.info("Proszę wgrać oba pliki po lewej stronie.")
    st.stop()

# ------------------------------
# 2) Parsowanie Excela zamówienia
# ------------------------------

df_order_raw = pd.read_excel(file_order, dtype=str)
syn_ean_ord = {normalize_col(c) for c in ["Symbol","symbol","kod ean","ean","kod produktu"]}
syn_qty_ord = {normalize_col(c) for c in ["ilość","ilosc","quantity","qty","sztuki"]}

ean_col = find_col(df_order_raw, syn_ean_ord)
qty_col = find_col(df_order_raw, syn_qty_ord)
if not ean_col or not qty_col:
    st.error(f"Brak kolumn EAN/Ilość w zamówieniu: {list(df_order_raw.columns)}")
    st.stop()

# Przygotowanie zamówienia
df_order = pd.DataFrame({
    "Symbol": df_order_raw[ean_col].astype(str).str.extract(r"(\d{13})")[0],
    "Zamówiona": df_order_raw[qty_col].astype(str)
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
ext = file_wz.name.lower().rsplit('.', 1)[-1]
rows = []

if ext == "pdf":
    with pdfplumber.open(file_wz) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables() or []
            for table in tables:
                if len(table) < 2:
                    continue
                hdr0, hdr1 = table[0], table[1]
                norm0 = [normalize_col(str(h)) for h in hdr0]
                norm1 = [normalize_col(str(h)) for h in hdr1]
                if any(c in syn_ean_wz for c in norm0) and any(c in syn_qty_wz for c in norm0):
                    header, data = hdr0, table[1:]
                elif any(c in syn_ean_wz for c in norm1) and any(c in syn_qty_wz for c in norm1):
                    header, data = hdr1, table[2:]
                else:
                    header, data = hdr0, table[1:]
                df_table = pd.DataFrame(data, columns=header)
                for _, r in df_table.iterrows():
                    # EAN
                    ean_match = next((m.group(1) for cell in r.astype(str) if (m := re.search(r"(\d{13})", cell))), None)
                    if not ean_match:
                        continue
                    # Ilość
                    qty_cell = None
                    for col in header:
                        if normalize_col(col) in syn_qty_wz:
                            qty_cell = r[col]
                            break
                    if qty_cell is None:
                        continue
                    val = str(qty_cell).replace(" ","").replace("\xa0","").replace(",",")").replace(",",")").replace(",",")").replace(",",")").replace(",",")").replace(",",")").replace(",",")").replace(",",")").replace(",",")").replace(",",")")).replace(",",".")
                    try:
                        qty = float(val)
                    except:
                        qty = 0.0
                    rows.append((ean_match, qty))
else:
    df_wz_raw = pd.read_excel(file_wz, dtype=str)
    ean_col_wz = find_col(df_wz_raw, syn_ean_wz)
    qty_col_wz = find_col(df_wz_raw, syn_qty_wz)
    if not ean_col_wz or not qty_col_wz:
        st.error(f"Brak kolumn EAN/Ilość w WZ: {list(df_wz_raw.columns)}")
        st.stop()
    for _, r in df_wz_raw.iterrows():
        if (m := re.search(r"(\d{13})", str(r[ean_col_wz]))):
            raw = str(r[qty_col_wz])
            val = raw.replace(" ","").replace("\xa0","").replace(",",".")
            try:
                qty = float(val)
            except:
                qty = 0.0
            rows.append((m.group(1), qty))

if not rows:
    st.error("Nie znaleziono danych w WZ.")
    st.stop()

# Suma WZ
df_wz = pd.DataFrame(rows, columns=["Symbol","Wydana"]).groupby("Symbol", as_index=False).sum()

# ------------------------------
# 4) Porównanie
# ------------------------------

df_cmp = pd.merge(df_order, df_wz, on="Symbol", how="outer", indicator=True)
=df_cmp[["Zamówiona","Wydana"]].fillna(0)
df_cmp["Różnica"] = df_cmp["Zamówiona"] - df_cmp["Wydana"]
status_map = {'left_only':'Brak we WZ','right_only':'Brak w zamówieniu'}

def get_status(row):
    return status_map.get(row._merge, 'OK' if row.Różnica==0 else 'Różni się')

df_cmp['Status'] = df_cmp.apply(get_status, axis=1)
df_cmp['Status'] = pd.Categorical(df_cmp['Status'], categories=['Różni się','Brak we WZ','Brak w zamówieniu','OK'], ordered=True)
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
    "⬇️ Pobierz raport",
    data=to_excel(df_cmp),
    file_name='porownanie_order_vs_wz.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)

# Podsumowanie
if (df_cmp['Status'] == 'OK').all():
    st.success("✅ Wszystkie pozycje się zgadzają")
else:
    st.error("❌ Wykryto rozbieżności w pozycjach")
