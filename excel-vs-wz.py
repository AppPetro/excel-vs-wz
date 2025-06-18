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
       - z PDF → przeprocesuje `extract_tables()` / z Excela bezpośrednio,
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
df_order_raw = pd.read_excel(uploaded_order, dtype=str)
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
    st.error(f"Brak kolumn EAN/Ilość w Excelu zamówienia: {list(df_order_raw.columns)}")
    st.stop()

df_order = pd.DataFrame({
    "Symbol": df_order_raw[col_ean_order].astype(str).str.extract(r"(\d{13})")[0],
    "Ilość": pd.to_numeric(df_order_raw[col_qty_order].str.replace(r"\s+","",regex=True), errors="coerce").fillna(0)
})

# 3) Parsowanie WZ
ext = uploaded_wz.name.lower().rsplit(".",1)[-1]

if ext == "pdf":
    with pdfplumber.open(uploaded_wz) as pdf:
        wz_rows = []
        def parse_wz_table(df_table: pd.DataFrame):
            for _, row in df_table.iterrows():
                # znajdź EAN 13-cyfrowy
                raw_ean = None
                for cell in row:
                    m = re.search(r"\b(\d{13})\b", str(cell))
                    if m:
                        raw_ean = m.group(1)
                        break
                if not raw_ean:
                    continue
                # znajdź ilość: ostatni występujący format liczba z przecinkiem lub kropką
                text = " ".join(map(str,row))
                qtys = re.findall(r"[\d\s]+[\.,]\d+", text)
                if qtys:
                    part = qtys[-1].replace(" ","").replace(",",".")
                    try:
                        qty = float(part)
                    except:
                        qty = 0.0
                else:
                    qty = 0.0
                wz_rows.append([raw_ean, qty])
        for page in pdf.pages:
            for table in page.extract_tables():
                if not table or len(table)<2:
                    continue
                df_page = pd.DataFrame(table[1:], columns=table[0])
                parse_wz_table(df_page)
    df_wz = pd.DataFrame(wz_rows, columns=["Symbol","Ilość_WZ"]).groupby("Symbol",as_index=False).sum()
else:
    df_wz_raw = pd.read_excel(uploaded_wz, dtype=str)
    df_wz = pd.DataFrame({
        "Symbol": df_wz_raw.apply(lambda r: re.search(r"(\d{13})", r.to_string()) and re.search(r"(\d{13})", r.to_string()).group(1), axis=1)
    })
    df_wz["Ilość_WZ"] = pd.to_numeric(df_wz_raw.apply(
        lambda r: next(iter(re.findall(r"[\d\s]+[\.,]\d+", r.to_string())),"0").replace(" ","").replace(",","."), axis=1
    ), errors="coerce").fillna(0)
    df_wz = df_wz.groupby("Symbol",as_index=False).sum()

# 4) Porównanie
df_cmp = pd.merge(df_order.groupby("Symbol",as_index=False).sum().rename(columns={"Ilość":"Zamówiona"}),
                  df_wz.rename(columns={"Ilość_WZ":"Wydana"}), on="Symbol", how="outer", indicator=True)
df_cmp.fillna({"Zamówiona":0,"Wydana":0}, inplace=True)
df_cmp["Różnica"] = df_cmp["Zamówiona"] - df_cmp["Wydana"]

# status
def stat(r):
    if r["_merge"]=="left_only": return "Brak we WZ"
    if r["_merge"]=="right_only": return "Brak w zamówieniu"
    return "OK" if r["Różnica"]==0 else "Różni się"

df_cmp["Status"] = df_cmp.apply(stat, axis=1)
order = ["Różni się","Brak we WZ","Brak w zamówieniu","OK"]
df_cmp["Status"] = pd.Categorical(df_cmp["Status"], categories=order, ordered=True)
df_cmp.sort_values(["Status","Symbol"], inplace=True)

# 5) Wyświetlenie i eksport
st.markdown("### Wynik porównania")
st.dataframe(df_cmp.style.format({"Zamówiona":"{:.0f}","Wydana":"{:.0f}","Różnica":"{:.0f}"}).apply(highlight_status_row,axis=1), use_container_width=True)

def to_excel(df):
    buf=BytesIO()
    writer=pd.ExcelWriter(buf,engine="openpyxl")
    df.to_excel(writer,index=False,sheet_name="Porównanie")
    writer.close()
    return buf.getvalue()

st.download_button("Pobierz raport",data=to_excel(df_cmp),file_name="raport.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# podsumowanie
if (df_cmp["Status"]=="OK").all():
    st.success("Wszystko się zgadza 🎉")
else:
    st.error("Są rozbieżności ❌")
