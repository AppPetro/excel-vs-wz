import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import re

st.set_page_config(
    page_title="📋 Porównywarka Zamówienie ↔ WZ (PDF→Excel)",
    layout="wide",
)

st.title("📋 Porównywarka Zamówienie ↔ WZ")

st.sidebar.header("🚀 Instrukcje")
st.sidebar.markdown(
    "1. Wgraj plik **Excel** z zamówieniem (kolumny EAN, Ilość).\n"
    "2. Wgraj plik **PDF** lub **Excel** z dokumentem WZ.\n"
    "3. Aplikacja zsumuje po **EAN** i porówna ilości."
)

# Funkcje pomocnicze
def normalize(text: str) -> str:
    return text.lower().replace(" ", "").replace("\xa0","").replace("_","")

def highlight(row):
    if row.Status == "OK": return ["background-color: #c6efce"]*len(row)
    return ["background-color: #ffc7ce"]*len(row)

# 1) >>> Wczytanie plików <<<
upload_order = st.sidebar.file_uploader("📥 Plik zamówienia (Excel)", type=["xlsx"])
upload_wz    = st.sidebar.file_uploader("📥 Plik WZ (PDF lub Excel)", type=["pdf","xlsx"])
if not upload_order or not upload_wz:
    st.info("Proszę wgrać oba pliki po lewej stronie.")
    st.stop()

# 2) >>> Excel zamówienia <<<
try:
    df_order_raw = pd.read_excel(upload_order, dtype=str)
except Exception as e:
    st.error(f"Błąd wczytywania Excela: {e}")
    st.stop()
# Znajdź kolumny EAN i Ilość
def find_column(df, patterns):
    for col in df.columns:
        if normalize(col) in patterns:
            return col
    return None

ean_ord = find_column(df_order_raw, {"symbol","kod ean","ean"})
qty_ord = find_column(df_order_raw, {"ilość","ilosc","qty","quantity"})
if not ean_ord or not qty_ord:
    st.error("Brak kolumn EAN lub Ilość w pliku zamówienia.")
    st.stop()
# Przygotuj df_order
order = pd.DataFrame()
order['Symbol'] = df_order_raw[ean_ord].astype(str).str.extract(r"(\d{13})")[0]
order['Zamówiona'] = pd.to_numeric(
    df_order_raw[qty_ord].astype(str).str.replace(r"[\s,]+", lambda m: m.group(0).strip().replace(',','.'), regex=True),
    errors='coerce'
).fillna(0)
order = order.groupby('Symbol', as_index=False).sum()

# 3) >>> Parsowanie WZ <<<
ext = upload_wz.name.lower().rsplit('.',1)[-1]
wz = pd.DataFrame(columns=['Symbol','Wydana'])
if ext == 'pdf':
    with pdfplumber.open(upload_wz) as pdf:
        rows = []
        for page in pdf.pages:
            tables = page.extract_tables() or []
            for table in tables:
                if len(table) < 2: continue
                df = pd.DataFrame(table[1:], columns=table[0])
                for _, r in df.iterrows():
                    text = ' '.join(map(str,r.values))
                    m_ean = re.search(r"\b(\d{13})\b", text)
                    if not m_ean: continue
                    ean = m_ean.group(1)
                    qtys = re.findall(r"[\d\s]+[\.,]\d+", text)
                    if not qtys: continue
                    num = qtys[-1].replace(' ','').replace(',','.')
                    try: val = float(num)
                    except: val = 0
                    rows.append((ean, val))
        if not rows:
            st.error("Nie znaleziono danych w PDF WZ.")
            st.stop()
        wz = pd.DataFrame(rows, columns=['Symbol','Wydana'])
else:
    tmp = pd.read_excel(upload_wz, dtype=str)
    rows = []
    for _, r in tmp.iterrows():
        line = ' '.join(r.astype(str))
        m = re.search(r"\b(\d{13})\b", line)
        if not m: continue
        ean = m.group(1)
        qtys = re.findall(r"[\d\s]+[\.,]\d+", line)
        if not qtys: continue
        num = qtys[-1].replace(' ','').replace(',','.')
        try: val = float(num)
        except: val = 0
        rows.append((ean, val))
    if not rows:
        st.error("Nie znaleziono danych w Excelu WZ.")
        st.stop()
    wz = pd.DataFrame(rows, columns=['Symbol','Wydana'])
# Zsumuj po EAN
df_wz = wz.groupby('Symbol', as_index=False).sum()

# 4) >>> Porównanie <<<
df_cmp = pd.merge(order, df_wz, on='Symbol', how='outer', indicator=True)
df_cmp[['Zamówiona','Wydana']] = df_cmp[['Zamówiona','Wydana']].fillna(0)
df_cmp['Różnica'] = df_cmp['Zamówiona'] - df_cmp['Wydana']
def get_status(m, diff):
    if m=='left_only': return 'Brak we WZ'
    if m=='right_only': return 'Brak w zamówieniu'
    return 'OK' if diff==0 else 'Różni się'
df_cmp['Status'] = df_cmp.apply(lambda x: get_status(x._merge, x.Różnica), axis=1)
# Sortowanie
df_cmp['Status'] = pd.Categorical(df_cmp['Status'], categories=['Różni się','Brak we WZ','Brak w zamówieniu','OK'], ordered=True)
df_cmp.sort_values(['Status','Symbol'], inplace=True)

# 5) >>> Wyświetlenie <<<
st.markdown("### 📊 Wyniki")
st.dataframe(df_cmp.style.format({
    'Zamówiona':'{:.0f}','Wydana':'{:.0f}','Różnica':'{:.0f}'
}).apply(highlight, axis=1), use_container_width=True)

# 6) >>> Eksport <<<
def to_xl(df):
    buf=BytesIO();w=pd.ExcelWriter(buf,engine='openpyxl');df.to_excel(w,index=False,sheet_name='Porównanie');w.close();return buf.getvalue()
st.download_button('⬇️ Pobierz raport', data=to_xl(df_cmp), file_name='porownanie.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# Podsumowanie
if (df_cmp['Status']=='OK').all(): st.success('✅ Wszystko się zgadza')
else: st.error('❌ Rozbieżności wykryte')
