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
       - EAN: `Symbol`, `symbol`, `kod ean`, `kod_ean`, `ean`, `kod produktu`, `kod_produktu`
       - Ilość: `Ilość`, `Ilosc`, `ilosc`, `Quantity`, `quantity`, `Qty`, `qty`, `sztuki`, `sztuka`
    2. Wgraj WZ w formie **PDF** (lub Excel), gdzie kolumna EAN może się nazywać:
       - `Kod produktu`, `kod produktu`, `kod_produktu`, `EAN`, `ean`, `symbol`
       - Kolumna ilości może nazywać się: `Ilość`, `Ilosc`, `ilosc`, `Quantity`, `quantity`, `Qty`, `qty`
    3. Aplikacja:
       - rozpozna synonimy kolumn w obu plikach,
       - z PDF → każdej stronie wyciągnie tabelę przez `extract_tables()` i „napasuje” kolumny EAN + Ilość (lub odtworzy rozbitą ilość),
       - zsumuje po EAN-ach i porówna z zamówieniem,
       - wyświetli wynik w tabeli, kolorując wiersze na zielono (OK) lub czerwono (gdy coś nie pasuje),
       - wyświetli komunikat u dołu w kolorze zielonym (“Pozycje się zgadzają”) lub czerwonym (“Pozycje się nie zgadzają”),
       - pozwoli pobrać raport jako Excel.
    """
)

# =============================================================================
# Funkcja do kolorowania wierszy wg kolumny "Status"
# =============================================================================
def highlight_status_row(row):
    """
    Zwraca listę stylów CSS dla jednego wiersza:
    - zielone tło, jeśli Status == "OK"
    - czerwone tło, w przeciwnym razie
    """
    color = "#c6efce" if row["Status"] == "OK" else "#ffc7ce"
    return [f"background-color: {color}" for _ in row.index]

# =============================================================================
# 1) Sidebar: wgrywanie plików
# =============================================================================
st.sidebar.header("Krok 1: Wgraj plik ZAMÓWIENIE (Excel)")
uploaded_order = st.sidebar.file_uploader(
    label="Wybierz plik Excel (zamówienie)",
    type=["xlsx"],
    key="order_uploader"
)

st.sidebar.header("Krok 2: Wgraj plik WZ (PDF lub Excel)")
uploaded_wz = st.sidebar.file_uploader(
    label="Wybierz plik WZ (PDF lub Excel)",
    type=["pdf", "xlsx"],
    key="wz_uploader"
)

st.sidebar.markdown(
    """
    - Dla **PDF**: parsujemy **wszystkie strony** przez `extract_tables()`,  
      rozpoznając synonimy kolumn EAN i ilości,  
      a gdy brak bezpośredniej kolumny „Ilość” → składamy z „Termin ważności Ilo” + „ść Waga brutto”.  
    - Dla **Excel (WZ→.xlsx)**: wczytujemy bezpośrednio kolumny `Kod produktu` (synonimy) + `Ilość` (synonimy).
    """
)

if uploaded_order is None or uploaded_wz is None:
    st.info("Proszę wgrać oba pliki: Excel (zamówienie) oraz PDF/Excel (WZ).")
    st.stop()

# =============================================================================
# 2) Wczytanie Excel z zamówieniem (synonimy kolumn)
# =============================================================================
try:
    df_order_raw = pd.read_excel(uploaded_order, dtype=str)
except Exception as e:
    st.error(f"Nie udało się wczytać pliku zamówienia:\n```{e}```")
    st.stop()

# Synonimy dla kolumny EAN w zamówieniu
synonyms_ean_order = {
    col.lower().replace(" ", "").replace("_", ""): col
    for col in [
        "Symbol", "symbol", "kod ean", "kod_ean",
        "ean", "kod produktu", "kod_produktu"
    ]
}
# Synonimy dla kolumny Ilość w zamówieniu
synonyms_qty_order = {
    col.lower().replace(" ", "").replace("_", ""): col
    for col in [
        "Ilość", "Ilosc", "ilosc", "Quantity",
        "quantity", "Qty", "qty", "sztuki", "sztuka"
    ]
}

def find_column_by_synonyms(df: pd.DataFrame, synonyms: dict):
    """
    Znajduje w df kolumnę, której uproszczona nazwa (bez spacji/underscore, małe litery)
    pasuje do któregoś klucza słownika synonyms.
    Zwraca oryginalną nazwę kolumny albo None.
    """
    for raw_col in df.columns:
        key = raw_col.lower().replace(" ", "").replace("_", "")
        if key in synonyms:
            return raw_col
    return None

col_ean_order = find_column_by_synonyms(df_order_raw, synonyms_ean_order)
col_qty_order = find_column_by_synonyms(df_order_raw, synonyms_qty_order)

if col_ean_order is None or col_qty_order is None:
    st.error(
        "Excel z zamówieniem musi zawierać kolumnę z EAN-em "
        "(np. `Symbol`, `kod ean`, `ean`) oraz kolumnę z ilością "
        "(np. `Ilość`, `quantity`, `qty`, `sztuki`).\n"
        f"Znalezione nagłówki: {list(df_order_raw.columns)}"
    )
    st.stop()

# Oczyszczanie i konwersja wartości
df_order = pd.DataFrame()
df_order["Symbol"] = (
    df_order_raw[col_ean_order].astype(str)
    .str.strip()
    .str.replace(r"\.0+$", "", regex=True)
)
df_order["Ilość"] = pd.to_numeric(df_order_raw[col_qty_order], errors="coerce").fillna(0)

# =============================================================================
# 3) Wczytanie WZ (PDF lub Excel), z synonimami kolumn
# =============================================================================
extension = uploaded_wz.name.lower().rsplit(".", maxsplit=1)[-1]

if extension == "pdf":
    try:
        with pdfplumber.open(uploaded_wz) as pdf:
            wz_rows = []

            # Synonimy dla kolumny EAN w WZ
            synonyms_ean_wz = {
                col.lower().replace(" ", "").replace("_", ""): col
                for col in [
                    "Kod produktu", "kod produkt", "kod_produktu",
                    "EAN", "ean", "symbol"
                ]
            }
            # Synonimy dla kolumny Ilość w WZ (bez rozbicia)
            synonyms_qty_wz = {
                col.lower().replace(" ", "").replace("_", ""): col
                for col in [
                    "Ilość", "Ilosc", "ilosc",
                    "Quantity", "quantity", "Qty", "qty"
                ]
            }

            def parse_wz_table(df_table: pd.DataFrame):
                """
                Zidentyfikuje w df_table:
                - kolumnę EAN (przez synonimy w synonyms_ean_wz)
                - kolumnę Ilość (przez synonimy w synonyms_qty_wz),
                  albo jeśli jej nie ma – rozbija ilość z
                  'Termin ważności Ilo' + 'ść Waga brutto'.
                Pomija wiersze, w których EAN != 13 cyfr.
                """
                cols = list(df_table.columns)

                # 1) Znajdź kolumnę EAN
                col_ean = None
                for raw_col in cols:
                    key = raw_col.lower().replace(" ", "").replace("_", "")
                    if key in synonyms_ean_wz:
                        col_ean = raw_col
                        break
                if col_ean is None:
                    return  # Brak kolumny EAN → pomiń całą tabelę

                # 2) Znajdź kolumnę Ilość (jeśli istnieje „prostą”)
                col_qty = None
                for raw_col in cols:
                    key = raw_col.lower().replace(" ", "").replace("_", "")
                    if key in synonyms_qty_wz:
                        col_qty = raw_col
                        break

                if col_qty:
                    # Mamy bezpośrednią kolumnę „Ilość”
                    for _, row in df_table.iterrows():
                        raw_ean_full = str(row[col_ean]).strip()
                        # Wyciągamy ostatni token (np. "1 4250231536748" → "4250231536748")
                        last_token = raw_ean_full.split()[-1] if raw_ean_full else ""
                        # Sprawdź, czy to dokładnie 13 cyfr
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
                    # Brak prostej kolumny „Ilość” → zakładamy „Termin ważności Ilo” + „ść Waga brutto”
                    col_part_int = next(
                        (c for c in cols if "termin" in c.lower() and "ilo" in c.lower()),
                        None
                    )
                    col_part_dec = next(
                        (c for c in cols if "waga" in c.lower()),
                        None
                    )
                    if col_part_int is None or col_part_dec is None:
                        return  # Niepoprawne nagłówki → pomiń

                    for _, row in df_table.iterrows():
                        raw_ean_full = str(row[col_ean]).strip()
                        last_token = raw_ean_full.split()[-1] if raw_ean_full else ""
                        if not re.fullmatch(r"\d{13}", last_token):
                            continue
                        ean = last_token

                        # --- Nowa, poprawiona metoda wyciągania ilości ---
                        # 3.1) Z 'Termin ważności Ilo' wyciągamy pierwszą liczbę przed przecinkiem:
                        part_int_cell = str(row[col_part_int])
                        m_int = re.search(r"(\d+),", part_int_cell)
                        raw_int = m_int.group(1) if m_int else "0"

                        # 3.2) Z 'ść Waga brutto' wyciągamy pierwsze dwie cyfry po przecinku:
                        part_dec_cell = str(row[col_part_dec])
                        m_dec = re.search(r",(\d{2})", part_dec_cell)
                        raw_dec = m_dec.group(1) if m_dec else "00"

                        # 3.3) Scalamy w formę "3,00" → 3.00
                        qty_str = f"{raw_int},{raw_dec}"
                        try:
                            qty = float(qty_str.replace(",", "."))
                        except:
                            qty = 0.0

                        wz_rows.append([ean, qty])

            # Przechodzimy po wszystkich stronach PDF-a
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table and len(table) > 1:
                        df_page = pd.DataFrame(table[1:], columns=table[0])
                        parse_wz_table(df_page)

    except Exception as e:
        st.error(f"Nie udało się przeczytać PDF przez pdfplumber:\n```{e}```")
        st.stop()

    if not wz_rows:
        st.error("Nie znaleziono żadnych danych w PDF WZ.")
        st.stop()

    # Tworzymy DataFrame i sumujemy wszystkie wiersze
    df_wz = pd.DataFrame(wz_rows, columns=["Symbol", "Ilość_WZ"])
    df_wz["Symbol"] = df_wz["Symbol"].astype(str).str.strip()
    df_wz["Ilość_WZ"] = pd.to_numeric(df_wz["Ilość_WZ"], errors="coerce").fillna(0)

else:
    # Jeżeli wgrano Excel z WZ
    try:
        df_wz_raw = pd.read_excel(uploaded_wz, dtype=str)
    except Exception as e:
        st.error(f"Nie udało się wczytać Excela WZ:\n```{e}```")
        st.stop()

    # Synonimy dla kolumny EAN w WZ
    synonyms_ean_wz = {
        col.lower().replace(" ", "").replace("_", ""): col
        for col in [
            "Kod produktu", "kod produkt", "kod_produktu",
            "EAN", "ean", "symbol"
        ]
    }
    # Synonimy dla kolumny Ilość w WZ
    synonyms_qty_wz = {
        col.lower().replace(" ", "").replace("_", ""): col
        for col in [
            "Ilość", "Ilosc", "ilosc",
            "Quantity", "quantity", "Qty", "qty"
        ]
    }

    col_ean_wz = None
    for raw_col in df_wz_raw.columns:
        key = raw_col.lower().replace(" ", "").replace("_", "")
        if key in synonyms_ean_wz:
            col_ean_wz = raw_col
            break

    col_qty_wz = None
    for raw_col in df_wz_raw.columns:
        key = raw_col.lower().replace(" ", "").replace("_", "")
        if key in synonyms_qty_wz:
            col_qty_wz = raw_col
            break

    if col_ean_wz is None or col_qty_wz is None:
        st.error(
            "Excel WZ musi zawierać kolumnę z EAN-em "
            "(np. `Kod produktu`, `EAN`, `symbol`) oraz kolumnę z ilością "
            "(np. `Ilość`, `quantity`, `qty`).\n"
            f"Znalezione nagłówki: {list(df_wz_raw.columns)}"
        )
        st.stop()

    # Dodatkowo sprawdźmy, czy ean to dokładnie 13 cyfr
    tmp_sym = df_wz_raw[col_ean_wz].astype(str).str.strip().str.split().str[-1]
    mask_valid_ean = tmp_sym.str.fullmatch(r"\d{13}")

    df_wz = pd.DataFrame({
        "Symbol": tmp_sym[mask_valid_ean],
        "Ilość_WZ": pd.to_numeric(
            df_wz_raw.loc[mask_valid_ean, col_qty_wz].astype(str)
            .str.replace(",", ".")
            .str.replace(r"\s+", "", regex=True),
            errors="coerce"
        ).fillna(0)
    })

# =============================================================================
# 4) Grupowanie i sumowanie
# =============================================================================
df_order_grouped = (
    df_order
    .groupby("Symbol", as_index=False)
    .agg({"Ilość": "sum"})
    .rename(columns={"Ilość": "Zamówiona_ilość"})
)

df_wz_grouped = (
    df_wz
    .groupby("Symbol", as_index=False)
    .agg({"Ilość_WZ": "sum"})
    .rename(columns={"Ilość_WZ": "Wydana_ilość"})
)

# =============================================================================
# 5) Merge + kolumna Status + Różnica
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
# 6) Wyświetlenie z kolorowaniem i eksport
# =============================================================================
st.markdown("### 📊 Wynik porównania")

styled = (
    df_compare
    .style
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
    label="⬇️ Pobierz raport jako Excel",
    data=to_excel(df_compare),
    file_name="porownanie_order_vs_wz.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# =============================================================================
# 7) Komunikat końcowy, kolorowany według statusu
# =============================================================================
all_ok = (df_compare["Status"] == "OK").all()

if all_ok:
    st.markdown(
        "<h4 style='color:green;'>✅ Pozycje się zgadzają</h4>",
        unsafe_allow_html=True
    )
else:
    st.markdown(
        "<h4 style='color:red;'>❌ Pozycje się nie zgadzają</h4>",
        unsafe_allow_html=True
    )
