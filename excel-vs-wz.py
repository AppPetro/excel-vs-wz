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
    1. Wgraj Excel z zamówieniem, zawierający kolumny z nazwami EAN i ilości (mogą to być synonimy):
       - EAN: `Symbol`, `symbol`, `kod ean`, `ean`, `kod produktu`
       - Ilość: `Ilość`, `Ilosc`, `Quantity`, `Qty`, `sztuki`
    2. Wgraj WZ w formie **PDF** (lub Excel), gdzie kolumna EAN może się nazywać:
       - `Kod produktu`, `EAN`, `symbol`
       - Kolumna ilości: `Ilość`, `Quantity`, `Qty`
    3. Aplikacja:
       - rozpozna synonimy kolumn,
       - z PDF → przeprocesuje strony przez `extract_tables()`,
       - zsumuje po EAN-ach i porówna z zamówieniem,
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
st.sidebar.header("Krok 1: Wgraj plik ZAMÓWIENIE (Excel)")
uploaded_order = st.sidebar.file_uploader("Wybierz plik Excel (zamówienie)", type=["xlsx"], key="order_uploader")
st.sidebar.header("Krok 2: Wgraj plik WZ (PDF lub Excel)")
uploaded_wz = st.sidebar.file_uploader("Wybierz plik WZ (PDF lub Excel)", type=["pdf", "xlsx"], key="wz_uploader")

if uploaded_order is None or uploaded_wz is None:
    st.info("Proszę wgrać oba pliki: Excel (zamówienie) oraz PDF/Excel (WZ).")
    st.stop()

# -------------------------
# 2) Parsowanie Zamówienia
# -------------------------
try:
    df_order_raw = pd.read_excel(uploaded_order, dtype=str)
except Exception as e:
    st.error(f"Nie udało się wczytać pliku zamówienia:\n```{e}```")
    st.stop()

synonyms_ean_order = { normalize_col_name(c): c for c in ["Symbol","symbol","kod ean","ean","kod produktu"] }
synonyms_qty_order = { normalize_col_name(c): c for c in ["Ilość","Ilosc","Quantity","Qty","sztuki"] }

def find_column(df, syns):
    for c in df.columns:
        if normalize_col_name(c) in syns:
            return c
    return None

col_ean_order = find_column(df_order_raw, synonyms_ean_order)
col_qty_order = find_column(df_order_raw, synonyms_qty_order)
if not col_ean_order or not col_qty_order:
    st.error(
        "Excel z zamówieniem musi zawierać kolumnę z EAN-em i kolumnę z ilością.\n"
        f"Znalezione nagłówki: {list(df_order_raw.columns)}"
    )
    st.stop()

df_order = pd.DataFrame({
    "Symbol": df_order_raw[col_ean_order].astype(str).str.strip().str.replace(r"\.0+$","",regex=True),
    "Ilość": pd.to_numeric(df_order_raw[col_qty_order], errors="coerce").fillna(0)
})

# -------------------------
# 3) Parsowanie WZ
# -------------------------
extension = uploaded_wz.name.lower().rsplit(".",1)[-1]
if extension == "pdf":
    try:
        with pdfplumber.open(uploaded_wz) as pdf:
            wz_rows = []

            synonyms_ean_wz = { normalize_col_name(c): c for c in ["Kod produktu","EAN","symbol"] }
            synonyms_qty_wz = { normalize_col_name(c): c for c in ["Ilość","Ilosc","Quantity","Qty"] }

            def parse_wz_table(df_table: pd.DataFrame):
                cols = list(df_table.columns)

                # 1) EAN
                col_ean = next((c for c in cols if normalize_col_name(c) in synonyms_ean_wz), None)
                if not col_ean:
                    return

                # 2) Ilość
                # strict match → prosta kolumna
                col_qty = next(
                    (c for c in cols if normalize_col_name(c) in synonyms_qty_wz),
                    None
                )

                if col_qty:
                    for _, row in df_table.iterrows():
                        raw_ean = str(row[col_ean]).strip().split()[-1]
                        if not re.fullmatch(r"\d{13}", raw_ean):
                            continue

                        raw_qty = str(row[col_qty]).strip().replace(",",".").replace(" ","")
                        try:
                            qty = float(raw_qty)
                        except:
                            qty = 0.0
                        wz_rows.append([raw_ean, qty])

                else:
                    # ostatnia kolumna jako fallback
                    for _, row in df_table.iterrows():
                        raw_ean = str(row[col_ean]).strip().split()[-1]
                        if not re.fullmatch(r"\d{13}", raw_ean):
                            continue
                        # bierzemy ostatnią kolumnę
                        raw_qty = str(row.iloc[-1]).strip().replace(",",".").replace(" ","")
                        try:
                            qty = float(raw_qty)
                        except:
                            qty = 0.0
                        wz_rows.append([raw_ean, qty])

            # ———> **TU jest patch: wybieramy odpowiedni wiersz nagłówka**
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if not table or len(table) < 2:
                        continue

                    header0 = table[0]
                    header1 = table[1]
                    norm0 = [normalize_col_name(str(c)) for c in header0]

                    # jeśli w pierwszym wierszu są synonimy EAN → to on jest headerem
                    if any(k in synonyms_ean_wz for k in norm0):
                        header = header0
                        data = table[1:]
                    else:
                        header = header1
                        data = table[2:]

                    if not data:
                        continue

                    df_page = pd.DataFrame(data, columns=header)
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

    synonyms_ean_wz = { normalize_col_name(c): c for c in ["Kod produktu","EAN","symbol"] }
    synonyms_qty_wz = { normalize_col_name(c): c for c in ["Ilość","Ilosc","Quantity","Qty"] }

    col_ean_wz = next((c for c in df_wz_raw.columns if normalize_col_name(c) in synonyms_ean_wz), None)
    col_qty_wz = next((c for c in df_wz_raw.columns if normalize_col_name(c) in synonyms_qty_wz), None)
    if not col_ean_wz or not col_qty_wz:
        st.error(
            "Excel WZ musi zawierać kolumnę z EAN-em i kolumnę z ilością.\n"
            f"Znalezione nagłówki: {list(df_wz_raw.columns)}"
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
                .str.replace(r"\s+","",regex=True),
            errors="coerce"
        ).fillna(0)
    })

# -------------------------
# 4) Grupowanie i sumowanie
# -------------------------
df_order_g = df_order.groupby("Symbol", as_index=False).agg({"Ilość":"sum"}).rename(columns={"Ilość":"Zamówiona_ilość"})
df_wz_g    = df_wz.groupby("Symbol", as_index=False).agg({"Ilość_WZ":"sum"}).rename(columns={"Ilość_WZ":"Wydana_ilość"})

# -------------------------
# 5) Porównanie
# -------------------------
df_cmp = pd.merge(df_order_g, df_wz_g, on="Symbol", how="outer", indicator=True)
df_cmp["Zamówiona_ilość"] = df_cmp["Zamówiona_ilość"].fillna(0)
df_cmp["Wydana_ilość"]    = df_cmp["Wydana_ilość"].fillna(0)
df_cmp["Różnica"]         = df_cmp["Zamówiona_ilość"] - df_cmp["Wydana_ilość"]

def status(r):
    if r["_merge"]=="left_only":  return "Brak we WZ"
    if r["_merge"]=="right_only": return "Brak w zamówieniu"
    return "OK" if r["Różnica"]==0 else "Różni się"

df_cmp["Status"] = df_cmp.apply(status,axis=1)
order_stat = ["Różni się","Brak we WZ","Brak w zamówieniu","OK"]
df_cmp["Status"] = pd.Categorical(df_cmp["Status"], categories=order_stat, ordered=True)
df_cmp = df_cmp.sort_values(["Status","Symbol"])

# -------------------------
# 6) Wyświetlenie i eksport
# -------------------------
st.markdown("### 📊 Wynik porównania")
styled = (
    df_cmp.style
          .format({"Zamówiona_ilość":"{:.0f}","Wydana_ilość":"{:.0f}","Różnica":"{:.0f}"})
          .apply(highlight_status_row, axis=1)
)
st.dataframe(styled, use_container_width=True)

def to_excel(df):
    out = BytesIO()
    w = pd.ExcelWriter(out, engine="openpyxl")
    df.to_excel(w, index=False, sheet_name="Porównanie")
    w.close()
    return out.getvalue()

st.download_button(
    "⬇️ Pobierz raport Excel",
    data=to_excel(df_cmp),
    file_name="porownanie_order_vs_wz.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

all_ok = (df_cmp["Status"]=="OK").all()
if all_ok:
    st.markdown("<h4 style='color:green;'>✅ Pozycje się zgadzają</h4>", unsafe_allow_html=True)
else:
    st.markdown("<h4 style='color:red;'>❌ Pozycje się nie zgadzają</h4>", unsafe_allow_html=True)
