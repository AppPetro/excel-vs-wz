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

# =============================================================================
# Funkcja do kolorowania wierszy wg kolumny "Status"
# =============================================================================
def highlight_status_row(row):
    color = "#c6efce" if row["Status"] == "OK" else "#ffc7ce"
    return [f"background-color: {color}" for _ in row.index]

# =============================================================================
# Uproszczanie nazwy kolumny
# =============================================================================
def normalize_col_name(name: str) -> str:
    return name.lower().replace(" ", "").replace("\xa0", "").replace("_", "")

# =============================================================================
# 1) Sidebar: wgrywanie plików
# =============================================================================
st.sidebar.header("Krok 1: Wgraj plik ZAMÓWIENIE (Excel)")
uploaded_order = st.sidebar.file_uploader("Wybierz plik Excel (zamówienie)", type=["xlsx"], key="order_uploader")

st.sidebar.header("Krok 2: Wgraj plik WZ (PDF lub Excel)")
uploaded_wz = st.sidebar.file_uploader("Wybierz plik WZ (PDF lub Excel)", type=["pdf", "xlsx"], key="wz_uploader")

if uploaded_order is None or uploaded_wz is None:
    st.info("Proszę wgrać oba pliki: Excel (zamówienie) oraz PDF/Excel (WZ).")
    st.stop()

# =============================================================================
# 2) Wczytanie Excel z zamówieniem
# =============================================================================
try:
    df_order_raw = pd.read_excel(uploaded_order, dtype=str)
except Exception as e:
    st.error(f"Nie udało się wczytać pliku zamówienia:\n```{e}```")
    st.stop()

# Synonimy EAN / Qty
synonyms_ean_order = {
    normalize_col_name(col): col
    for col in ["Symbol", "symbol", "kod ean", "ean", "kod produktu"]
}
synonyms_qty_order = {
    normalize_col_name(col): col
    for col in ["Ilość", "Ilosc", "Quantity", "Qty", "sztuki"]
}

def find_column_by_synonyms(df: pd.DataFrame, synonyms: dict):
    for raw_col in df.columns:
        if normalize_col_name(raw_col) in synonyms:
            return raw_col
    return None

col_ean_order = find_column_by_synonyms(df_order_raw, synonyms_ean_order)
col_qty_order = find_column_by_synonyms(df_order_raw, synonyms_qty_order)
if col_ean_order is None or col_qty_order is None:
    st.error(
        "Excel z zamówieniem musi zawierać kolumnę z EAN-em i kolumnę z ilością.\n"
        f"Znalezione nagłówki: {list(df_order_raw.columns)}"
    )
    st.stop()

df_order = pd.DataFrame({
    "Symbol": df_order_raw[col_ean_order].astype(str).str.strip().str.replace(r"\.0+$", "", regex=True),
    "Ilość": pd.to_numeric(df_order_raw[col_qty_order], errors="coerce").fillna(0)
})

# =============================================================================
# 3) Wczytanie WZ
# =============================================================================
extension = uploaded_wz.name.lower().rsplit(".", 1)[-1]
if extension == "pdf":
    try:
        with pdfplumber.open(uploaded_wz) as pdf:
            wz_rows = []

            synonyms_ean_wz = {
                normalize_col_name(col): col
                for col in ["Kod produktu", "EAN", "symbol"]
            }
            synonyms_qty_wz = {
                normalize_col_name(col): col
                for col in ["Ilość", "Ilosc", "Quantity", "Qty"]
            }

            def parse_wz_table(df_table: pd.DataFrame):
                cols = list(df_table.columns)

                # 1) EAN
                col_ean = None
                for raw_col in cols:
                    if normalize_col_name(raw_col) in synonyms_ean_wz:
                        col_ean = raw_col
                        break
                if col_ean is None:
                    return

                # 2) Ilość – direct match lub substring match
                col_qty = None
                for raw_col in cols:
                    key = normalize_col_name(raw_col)
                    if any(syn in key for syn in synonyms_qty_wz):
                        col_qty = raw_col
                        break

                if col_qty:
                    # prosta kolumna Ilość
                    for _, row in df_table.iterrows():
                        raw_ean_full = str(row[col_ean]).strip()
                        last_token = raw_ean_full.split()[-1] if raw_ean_full else ""
                        if not re.fullmatch(r"\d{13}", last_token):
                            continue
                        ean = last_token

                        raw_qty = str(row[col_qty]).strip().replace(",", ".").replace(" ", "")
                        try:
                            qty = float(raw_qty)
                        except:
                            qty = 0.0
                        wz_rows.append([ean, qty])

                else:
                    # broken header: Termin ważności Ilość + Waga brutto
                    col_part_int = None
                    col_part_dec = None
                    for raw_col in cols:
                        low = normalize_col_name(raw_col)
                        if "termin" in low and "ilo" in low:
                            col_part_int = raw_col
                        if "waga" in low:
                            col_part_dec = raw_col
                    if col_part_int is None or col_part_dec is None:
                        return

                    for _, row in df_table.iterrows():
                        raw_ean_full = str(row[col_ean]).strip()
                        last_token = raw_ean_full.split()[-1] if raw_ean_full else ""
                        if not re.fullmatch(r"\d{13}", last_token):
                            continue
                        ean = last_token

                        # integer part from 'Termin ważności Ilo'
                        part_int_cell = str(row[col_part_int])
                        m_int = re.search(r"(\d+),?$", part_int_cell)
                        raw_int = m_int.group(1) if m_int else "0"

                        # ignorujemy część dziesiętną (waga brutto!)
                        raw_dec = "00"

                        qty_str = f"{raw_int},{raw_dec}"
                        try:
                            qty = float(qty_str.replace(",", "."))
                        except:
                            qty = 0.0
                        wz_rows.append([ean, qty])

            # iterujemy po stronach PDF
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if not table or len(table) < 2:
                        continue
                    df_page = pd.DataFrame(table[1:], columns=table[0])
                    parse_wz_table(df_page)

    except Exception as e:
        st.error(f"Nie udało się przetworzyć PDF:\n```{e}```")
        st.stop()

    if not wz_rows:
        st.error("Nie znaleziono żadnych danych w PDF WZ.")
        st.stop()

    df_wz = pd.DataFrame(wz_rows, columns=["Symbol", "Ilość_WZ"])
    df_wz["Symbol"] = df_wz["Symbol"].astype(str).str.strip()
    df_wz["Ilość_WZ"] = pd.to_numeric(df_wz["Ilość_WZ"], errors="coerce").fillna(0)

else:
    # Excel z WZ
    try:
        df_wz_raw = pd.read_excel(uploaded_wz, dtype=str)
    except Exception as e:
        st.error(f"Nie udało się wczytać Excela WZ:\n```{e}```")
        st.stop()

    synonyms_ean_wz = {
        normalize_col_name(col): col
        for col in ["Kod produktu", "EAN", "symbol"]
    }
    synonyms_qty_wz = {
        normalize_col_name(col): col
        for col in ["Ilość", "Ilosc", "Quantity", "Qty"]
    }

    col_ean_wz = None
    for raw_col in df_wz_raw.columns:
        if normalize_col_name(raw_col) in synonyms_ean_wz:
            col_ean_wz = raw_col
            break
    col_qty_wz = None
    for raw_col in df_wz_raw.columns:
        if normalize_col_name(raw_col) in synonyms_qty_wz:
            col_qty_wz = raw_col
            break

    if col_ean_wz is None or col_qty_wz is None:
        st.error(
            "Excel WZ musi zawierać kolumnę z EAN-em i kolumnę z ilością.\n"
            f"Znalezione nagłówki: {list(df_wz_raw.columns)}"
        )
        st.stop()

    tmp_sym = df_wz_raw[col_ean_wz].astype(str).str.strip().str.split().str[-1]
    mask_valid = tmp_sym.str.fullmatch(r"\d{13}")
    df_wz = pd.DataFrame({
        "Symbol": tmp_sym[mask_valid],
        "Ilość_WZ": pd.to_numeric(
            df_wz_raw.loc[mask_valid, col_qty_wz]
                .astype(str)
                .str.replace(",", ".")
                .str.replace(r"\s+", "", regex=True),
            errors="coerce"
        ).fillna(0)
    })

# =============================================================================
# 4) Grupowanie i sumowanie
# =============================================================================
df_order_grouped = (
    df_order.groupby("Symbol", as_index=False)
            .agg({"Ilość": "sum"})
            .rename(columns={"Ilość": "Zamówiona_ilość"})
)
df_wz_grouped = (
    df_wz.groupby("Symbol", as_index=False)
          .agg({"Ilość_WZ": "sum"})
          .rename(columns={"Ilość_WZ": "Wydana_ilość"})
)

# =============================================================================
# 5) Porównanie
# =============================================================================
df_compare = pd.merge(
    df_order_grouped,
    df_wz_grouped,
    on="Symbol",
    how="outer",
    indicator=True
)
df_compare["Zamówiona_ilość"] = df_compare["Zamówiona_ilość"].fillna(0)
df_compare["Wydana_ilość"]    = df_compare["Wydana_ilość"].fillna(0)
df_compare["Różnica"] = df_compare["Zamówiona_ilość"] - df_compare["Wydana_ilość"]

def status_row(row):
    if row["_merge"] == "left_only":
        return "Brak we WZ"
    elif row["_merge"] == "right_only":
        return "Brak w zamówieniu"
    elif row["Zamówiona_ilość"] == row["Wydana_ilość"]:
        return "OK"
    else:
        return "Różni się"

df_compare["Status"] = df_compare.apply(status_row, axis=1)
order_status = ["Różni się", "Brak we WZ", "Brak w zamówieniu", "OK"]
df_compare["Status"] = pd.Categorical(df_compare["Status"], categories=order_status, ordered=True)
df_compare = df_compare.sort_values(["Status", "Symbol"])

# =============================================================================
# 6) Wyświetlenie i eksport
# =============================================================================
st.markdown("### 📊 Wynik porównania")
styled = (
    df_compare.style
              .format({
                  "Zamówiona_ilość": "{:.0f}",
                  "Wydana_ilość": "{:.0f}",
                  "Różnica": "{:.0f}"
              })
              .apply(highlight_status_row, axis=1)
)
st.dataframe(styled, use_container_width=True)

def to_excel(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine="openpyxl")
    df.to_excel(writer, index=False, sheet_name="Porównanie")
    writer.close()
    return output.getvalue()

st.download_button(
    label="⬇️ Pobierz raport Excel",
    data=to_excel(df_compare),
    file_name="porownanie_order_vs_wz.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# =============================================================================
# 7) Komunikat końcowy
# =============================================================================
all_ok = (df_compare["Status"] == "OK").all()
if all_ok:
    st.markdown("<h4 style='color:green;'>✅ Pozycje się zgadzają</h4>", unsafe_allow_html=True)
else:
    st.markdown("<h4 style='color:red;'>❌ Pozycje się nie zgadzają</h4>", unsafe_allow_html=True)
