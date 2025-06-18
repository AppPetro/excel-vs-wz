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

# -----------------------
# Funkcje pomocnicze
# -----------------------
def normalize_col(name: str) -> str:
    return re.sub(r"[\s\xa0_]+", "", name).lower()

def highlight_status(row):
    color = "#c6efce" if row.Status == "OK" else "#ffc7ce"
    return [f"background-color: {color}"] * len(row)

# -----------------------
# 1) Wczytanie plików
# -----------------------
st.sidebar.header("1️⃣ Wczytaj pliki")
file_order = st.sidebar.file_uploader("Excel zamówienia", type=["xlsx"])
file_wz    = st.sidebar.file_uploader("WZ (PDF lub Excel)", type=["pdf","xlsx"])
if not file_order or not file_wz:
    st.info("Proszę wgrać oba pliki: Zamówienie oraz WZ.")
    st.stop()

# -----------------------
# 2) Parsowanie zamówienia
# -----------------------
try:
    df_order_raw = pd.read_excel(file_order, dtype=str)
except Exception as e:
    st.error(f"Nie udało się wczytać pliku zamówienia: {e}")
    st.stop()
# Znalezienie kolumn EAN i Ilość
syn_ean = {normalize_col(c): c for c in ["Symbol","symbol","kod ean","ean","kod produktu"]}
syn_qty = {normalize_col(c): c for c in ["ilość","ilosc","quantity","qty","sztuki"]}
def find_col(df, syn):
    for c in df.columns:
        if normalize_col(c) in syn:
            return c
    return None
col_ean = find_col(df_order_raw, syn_ean)
col_qty = find_col(df_order_raw, syn_qty)
if not col_ean or not col_qty:
    st.error(f"Brak kolumn EAN lub Ilość w zamówieniu. Znalezione: {list(df_order_raw.columns)}")
    st.stop()
# Przygotowanie df_order
df_order = pd.DataFrame({
    "Symbol": df_order_raw[col_ean].astype(str).str.extract(r"(\d{13})")[0],
    "Zamówiona": df_order_raw[col_qty].astype(str)
        .str.replace(r"[\s\.]+", "", regex=True)  # usuń separatory tysięcy i kropki
        .str.replace(",", ".")
})
df_order["Zamówiona"] = pd.to_numeric(df_order["Zamówiona"], errors="coerce").fillna(0)
df_order = df_order.groupby("Symbol", as_index=False).sum()

# -----------------------
# 3) Parsowanie WZ
# -----------------------
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
                data = table[1:]
                df = pd.DataFrame(data, columns=header)
                for _, r in df.iterrows():
                    row_text = " ".join(map(str, r.values))
                    # EAN
                    m_ean = re.search(r"\b(\d{13})\b", row_text)
                    if not m_ean:
                        continue
                    ean = m_ean.group(1)
                    # Ilość: ostatnia liczba z przecinkiem/kropką
                    qtys = re.findall(r"[\d\s]+[\.,]\d+", row_text)
                    if not qtys:
                        continue
                    val = qtys[-1].replace(" ", "").replace(",", ".")
                    try:
                        qty = float(val)
                    except:
                        qty = 0.0
                    rows.append((ean, qty))
else:
    try:
        df_wz_raw = pd.read_excel(file_wz, dtype=str)
    except Exception as e:
        st.error(f"Nie udało się wczytać Excela WZ: {e}")
        st.stop()
    for _, r in df_wz_raw.iterrows():
        line = " ".join(r.astype(str))
        m_ean = re.search(r"\b(\d{13})\b", line)
        if not m_ean:
            continue
        ean = m_ean.group(1)
        qtys = re.findall(r"[\d\s]+[\.,]\d+", line)
        if not qtys:
            continue
        val = qtys[-1].replace(" ", "").replace(",", ".")
        try:
            qty = float(val)
        except:
            qty = 0.0
        rows.append((ean, qty))
# Tworzenie df_wz
# DEBUG: sprawdzenie wierszy dla EAN 4250231542008
target_ean = "4250231542008"
specific_rows = [qty for ean, qty in rows if ean == target_ean]
st.write(f"DEBUG: wiersze dla {target_ean}: {specific_rows} (suma: {sum(specific_rows)})")

df_wz = pd.DataFrame(rows, columns=["Symbol","Wydana"]).groupby("Symbol", as_index=False).sum()
if df_wz.empty:
    st.error("Brak danych wyciągniętych z WZ.")
    st.stop()

# -----------------------
# 4) Porównanie
# -----------------------
df_cmp = pd.merge(df_order, df_wz, on="Symbol", how="outer", indicator=True)
df_cmp[["Zamówiona","Wydana"]] = df_cmp[["Zamówiona","Wydana"]].fillna(0)
df_cmp["Różnica"] = df_cmp["Zamówiona"] - df_cmp["Wydana"]
def status(x):
    if x._merge == 'left_only': return 'Brak we WZ'
    if x._merge == 'right_only': return 'Brak w zamówieniu'
    return 'OK' if x.Różnica == 0 else 'Różni się'
df_cmp['Status'] = df_cmp.apply(status, axis=1)
order_cat = ['Różni się','Brak we WZ','Brak w zamówieniu','OK']
df_cmp['Status'] = pd.Categorical(df_cmp['Status'], categories=order_cat, ordered=True)
df_cmp.sort_values(['Status','Symbol'], inplace=True)

# -----------------------
# 5) Wyświetlenie i eksport
# -----------------------
st.markdown("### 📊 Wyniki porównania")
st.dataframe(
    df_cmp.style
        .format({'Zamówiona':'{:.0f}','Wydana':'{:.0f}','Różnica':'{:.0f}'})
        .apply(highlight_status, axis=1),
    use_container_width=True
)

def to_excel(dataframe):
    buf = BytesIO()
    writer = pd.ExcelWriter(buf, engine='openpyxl')
    dataframe.to_excel(writer, index=False, sheet_name='Porównanie')
    writer.close()
    return buf.getvalue()

st.download_button(
    label='⬇️ Pobierz raport',
    data=to_excel(df_cmp),
    file_name='porownanie_order_vs_wz.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)

# Podsumowanie
if (df_cmp['Status'] == 'OK').all():
    st.success("✅ Wszystkie pozycje się zgadzają")
else:
    st.error("❌ Wykryto rozbieżności w pozycjach")
