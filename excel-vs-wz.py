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
    1. Wgraj Excel z zamówieniem, zawierający kolumny:
       - `Symbol` (EAN, 13 cyfr)
       - `Ilość` (liczba sztuk)
    2. Wgraj WZ w formie **PDF** (lub Excel), zawierający:
       - `Kod produktu` (EAN)
       - `Ilość` (wydane sztuki)
       LUB w PDF-ie:
       - `Termin ważności Ilo` (część całkowita + data)
       - `ść Waga brutto`  (część dziesiętna + waga)
    3. Aplikacja:
       - wyciągnie **wszystkie** strony PDF przez `extract_tables()`,
       - z każdej tabeli odczyta EAN + Ilość (do prostych wierszy lub rozbitych kolumn),
       - zsumuje po EAN-ach, stworzy tabelę porównawczą i pozwoli pobrać raport.
    """
)

# ==============================
# 1) Sidebar: Wgrywanie plików
# ==============================
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
    - Dla **PDF**: użyjemy `pdfplumber.extract_tables()` na każdej stronie i 
      zawsze wykonamy spójne parsowanie EAN + Ilość,  
      bez rozbijania osobno pierwszej strony.  
    - Dla **Excel (WZ→.xlsx)**: wczytamy bezpośrednio kolumny `Kod produktu` + `Ilość`.
    """
)

if uploaded_order is None or uploaded_wz is None:
    st.info("Proszę wgrać oba pliki: Excel (zamówienie) oraz PDF/Excel (WZ).")
    st.stop()

# ==============================
# 2) Wczytanie Excel z zamówieniem
# ==============================
try:
    df_order = pd.read_excel(uploaded_order, dtype={"Symbol": str})
except Exception as e:
    st.error(f"Nie udało się wczytać pliku zamówienia:\n```{e}```")
    st.stop()

# Sprawdzenie poprawności nagłówków
if "Symbol" not in df_order.columns or "Ilość" not in df_order.columns:
    st.error(
        "Excel z zamówieniem musi mieć kolumny:\n"
        "- `Symbol` (EAN, 13 cyfr)\n"
        "- `Ilość` (liczba sztuk)\n\n"
        "Proszę sprawdzić dokładne nazwy nagłówków."
    )
    st.stop()

# Oczyszczenie EAN i konwersja ilości na numeric
df_order["Symbol"] = (
    df_order["Symbol"]
    .astype(str)
    .str.strip()
    .str.replace(r"\.0+$", "", regex=True)
)
df_order["Ilość"] = pd.to_numeric(df_order["Ilość"], errors="coerce").fillna(0)

# ==============================
# 3) Wczytanie WZ (PDF lub Excel)
# ==============================
extension = uploaded_wz.name.lower().rsplit(".", maxsplit=1)[-1]

if extension == "pdf":
    try:
        with pdfplumber.open(uploaded_wz) as pdf:
            # Lista tupli (EAN, ilość) z każdej strony
            wz_rows = []

            # Funkcja pomocnicza do parsowania pojedynczej tabeli w postaci DataFrame
            def parse_wz_table(df_table: pd.DataFrame):
                """
                Sprawdza nagłówki: 
                - jeśli jest kolumna "Ilość" → bierze ją bezpośrednio,
                - inaczej zakłada, że kolumny to 'Termin ważności Ilo' + 'ść Waga brutto'.
                Zwraca listę [ [ean, qty], ... ].
                """
                cols = [c.strip() for c in df_table.columns]

                # SPRóbuj znaleźć bezpośrednią kolumnę "Ilość"
                col_qty = next((c for c in cols if c.lower() == "ilość"), None)
                col_ean = next((c for c in cols if "kod" in c.lower() and "produkt" in c.lower()), None)

                if col_qty and col_ean:
                    for _, row in df_table.iterrows():
                        raw_ean = str(row[col_ean]).strip()
                        if raw_ean == "" or pd.isna(raw_ean):
                            continue
                        # Weź ostatni token EAN (czasem występuje prefiks "1 425023...")
                        ean = raw_ean.split()[-1]

                        raw_qty = str(row[col_qty]).strip().replace(",", ".").replace(" ", "")
                        try:
                            qty = float(raw_qty)
                        except:
                            qty = 0.0
                        wz_rows.append([ean, qty])

                else:
                    # Rozbita kolumna: 'Termin ważności Ilo' (data + część całk.) oraz 'ść Waga brutto' (część dzies.)
                    col_part_int = next((c for c in cols if "termin" in c.lower() and "ilo" in c.lower()), None)
                    col_part_dec = next((c for c in cols if "waga" in c.lower()), None)

                    if col_part_int is None or col_part_dec is None or col_ean is None:
                        # Niepoprawne nagłówki, ignorujemy tabelę
                        return

                    for _, row in df_table.iterrows():
                        raw_ean = str(row[col_ean]).strip()
                        if raw_ean == "" or pd.isna(raw_ean):
                            continue
                        ean = raw_ean.split()[-1]

                        # Część całkowita: np. "2027-11-27 60," → "60"
                        part_int_cell = str(row[col_part_int]).strip()
                        tokens_int = part_int_cell.split()
                        if len(tokens_int) >= 2:
                            raw_int = tokens_int[-1].replace(",", "").strip()
                        else:
                            raw_int = "0"

                        # Część dziesiętna: np. "00 15,00" → weź "00"
                        part_dec_cell = str(row[col_part_dec]).strip()
                        tokens_dec = part_dec_cell.split()
                        dec_part = tokens_dec[0].replace(".", "").strip() if tokens_dec else "00"

                        # Scal w "60,00"
                        if dec_part.startswith(","):
                            qty_str = f"{raw_int}{dec_part}"
                        else:
                            qty_str = f"{raw_int},{dec_part}"
                        try:
                            qty = float(qty_str.replace(",", "."))
                        except:
                            qty = 0.0

                        wz_rows.append([ean, qty])

            # Przechodzimy przez wszystkie strony PDF
            for page in pdf.pages:
                # extract_tables() → lista wszystkich tabel na stronie
                tables = page.extract_tables()
                for table in tables:
                    if table and len(table) > 1:
                        # Konwertuj listę list na DataFrame
                        df_page = pd.DataFrame(table[1:], columns=table[0])
                        parse_wz_table(df_page)

    except Exception as e:
        st.error(f"Nie udało się wczytać PDF-a przez pdfplumber:\n```{e}```")
        st.stop()

    if not wz_rows:
        st.error("Nie znaleziono żadnej tabeli WZ w PDF-ie.")
        st.stop()

    # Przemień listę [ [ean, qty], ... ] w DataFrame i zsumuj po EAN
    df_wz = pd.DataFrame(wz_rows, columns=["Symbol", "Ilość_WZ"])
    df_wz["Symbol"] = df_wz["Symbol"].astype(str).str.strip()
    df_wz["Ilość_WZ"] = pd.to_numeric(df_wz["Ilość_WZ"], errors="coerce").fillna(0)

else:
    # Jeżeli użytkownik wgrał gotowy Excel z WZ
    try:
        df_wz_raw = pd.read_excel(uploaded_wz, dtype={"Kod produktu": str})
    except Exception as e:
        st.error(f"Nie udało się wczytać Excela WZ:\n```{e}```")
        st.stop()

    if "Kod produktu" not in df_wz_raw.columns or "Ilość" not in df_wz_raw.columns:
        st.error(
            "Excel WZ musi mieć kolumny:\n"
            "- `Kod produktu` (EAN)\n"
            "- `Ilość` (wydane sztuki)\n\n"
            f"Znalezione nagłówki: {list(df_wz_raw.columns)}"
        )
        st.stop()

    df_wz = df_wz_raw.rename(columns={"Kod produktu": "Symbol", "Ilość": "Ilość_WZ"})
    df_wz["Symbol"] = df_wz["Symbol"].astype(str).str.strip().str.replace(r"\.0+$", "", regex=True)
    df_wz["Ilość_WZ"] = (
        df_wz["Ilość_WZ"].astype(str)
        .str.replace(",", ".", regex=False)
        .str.replace(r"\s+", "", regex=True)
    )
    df_wz["Ilość_WZ"] = pd.to_numeric(df_wz["Ilość_WZ"], errors="coerce").fillna(0)

# ==============================
# 4) Grupowanie i sumowanie
# ==============================
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

# ==============================
# 5) Merge + Status + Różnica
# ==============================
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

# Posortuj w kolejności: różni się, brak we WZ, brak w zamówieniu, OK
order_status = ["Różni się", "Brak we WZ", "Brak w zamówieniu", "OK"]
df_compare["Status"] = pd.Categorical(df_compare["Status"], categories=order_status, ordered=True)
df_compare = df_compare.sort_values(["Status", "Symbol"])

# ==============================
# 6) Wyświetlenie i eksport
# ==============================
st.markdown("### 📊 Wynik porównania")
st.dataframe(
    df_compare.style.format({
        "Zamówiona_ilość": "{:.0f}",
        "Wydana_ilość": "{:.0f}",
        "Różnica": "{:.0f}"
    }),
    use_container_width=True
)

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

st.success("✅ Gotowe! Porównanie wykonane pomyślnie.")
