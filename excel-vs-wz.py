import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO

st.set_page_config(
    page_title="📋 Porównywarka Zamówienie ↔ WZ (PDF→Excel)",
    layout="wide",
)

st.title("📋 Porównywarka Zamówienie (Excel) vs. WZ (PDF lub Excel)")

st.markdown(
    """
    **Instrukcja:**
    1. Wgraj plik Excel z zamówieniem, zawierający przynajmniej kolumny:
       - `Symbol` (EAN, np. 5029040012281)
       - `Ilość` (liczba zamawianych sztuk)
    2. Wgraj plik WZ w formacie **PDF** (lub, jeśli wolisz, gotowy Excel z WZ), 
       zawierający przynajmniej kolumny:
       - `Kod produktu` (EAN)
       - `Ilość` (liczba wydanych sztuk w danym wierszu WZ)
       LUB, w przypadku PDF-ów, nagłówek w stylu rozbitym na dwie kolumny:
       - `Termin ważności Ilo` + `ść Waga brutto`
    3. Aplikacja automatycznie:
       - wyciągnie wszystkie tabele z PDF-a (wszystkie strony) za pomocą `pdfplumber.extract_tables()`,
       - wykryje, czy kolumna „Ilość” jest zapisana wprost, czy rozbita na dwa pola (część całkowita i część dziesiętna),
       - zbuduje prawidłową wartość „Ilość”,
       - zsumuje te ilości po EAN-ach,
       - porówna z zamówieniem i wygeneruje raport z różnicą.
    """
)

# =====================================
# SIDEBAR: Upload plików
# =====================================
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
    - Jeśli wgrasz **PDF**, aplikacja użyje `pdfplumber` do wyciągnięcia WSZYSTKICH tabel z każdej strony, 
      następnie rozpozna, czy kolumna „Ilość” jest podana wprost, czy rozbita na dwie („Termin ważności Ilo” + „ść Waga brutto”).  
    - Jeśli wgrasz **Excel** (plik otrzymany z WZ→.xlsx), aplikacja wczyta kolumny `Kod produktu` i `Ilość` bezpośrednio.
    """
)

if uploaded_order is None or uploaded_wz is None:
    st.info("Proszę wgrać oba pliki po lewej stronie: plik z Zamówieniem oraz plik z WZ.  ")
    st.stop()

# =====================================
# 1) Wczytanie i przygotowanie Excela z ZAMÓWIENIEM
# =====================================
try:
    df_order = pd.read_excel(uploaded_order, dtype={"Symbol": str})
except Exception as e:
    st.error(f"Nie udało się wczytać pliku zamówienia:\n```{e}```")
    st.stop()

# Sprawdzenie, czy kolumny istnieją:
if "Symbol" not in df_order.columns or "Ilość" not in df_order.columns:
    st.error(
        "Plik ZAMÓWIENIE musi mieć kolumny:\n"
        "- `Symbol` (EAN)\n"
        "- `Ilość` (liczba zamawianych sztuk)\n\n"
        "Sprawdź, czy nagłówki dokładnie tak się nazywają (wielkość liter i spacje)."
    )
    st.stop()

# Oczyść EAN-y (usuń spacje i sufiks .0) i konwertuj zamówioną ilość na liczbę
df_order["Symbol"] = (
    df_order["Symbol"]
    .astype(str)
    .str.strip()
    .str.replace(r"\.0+$", "", regex=True)
)
df_order["Ilość"] = pd.to_numeric(df_order["Ilość"], errors="coerce").fillna(0)

# =====================================
# 2) Wczytanie i przygotowanie danych z WZ (PDF lub Excel)
# =====================================
file_ext = uploaded_wz.name.lower().rsplit(".", maxsplit=1)[-1]

if file_ext == "pdf":
    # ---------------------------------
    # 2a) Ekstrakcja WSZYSTKICH tabel z PDF za pomocą pdfplumber
    # ---------------------------------
    try:
        with pdfplumber.open(uploaded_wz) as pdf:
            all_tables = []
            for page in pdf.pages:
                tables_on_page = page.extract_tables()
                for table in tables_on_page:
                    # Jeśli tabela nie jest pusta i ma co najmniej 2 wiersze (nagłówek + dane)
                    if table and len(table) > 1:
                        df_page = pd.DataFrame(table[1:], columns=table[0])
                        all_tables.append(df_page)
    except Exception as e:
        st.error(f"Nie udało się przeczytać PDF-a przez pdfplumber:\n```{e}```")
        st.stop()

    if len(all_tables) == 0:
        st.error("Nie znaleziono żadnych tabel w pliku PDF WZ.")
        st.stop()

    # Scal wszystkie fragmenty tabel w jeden DataFrame
    df_wz_raw = pd.concat(all_tables, ignore_index=True)

    # Sprawdź nagłówki w df_wz_raw
    cols = list(df_wz_raw.columns)

    # 2a.1) Wariant A: jeśli istnieje kolumna nazwana dokładnie "Ilość"
    ilo_exists = next((col for col in cols if col.lower().strip() == "ilość"), None)
    if ilo_exists is not None:
        # Kolumna 'Ilość' jest dostępna od razu
        col_qty = ilo_exists
        col_ean = next((col for col in cols if "kod" in col.lower() and "produkt" in col.lower()), None)
        if col_ean is None:
            st.error(
                "Po scaleniu tabel z PDF nie znaleziono kolumny `Kod produktu`.\n"
                f"Znalezione nagłówki: {cols}"
            )
            st.stop()

        # Wyciągnij tylko te dwie kolumny
        df_wz = pd.DataFrame({
            "Symbol": df_wz_raw[col_ean].astype(str),
            "Ilość_WZ": df_wz_raw[col_qty]
        })

        # Oczyść EAN-y i skonwertuj 'Ilość_WZ' na float
        df_wz["Symbol"] = (
            df_wz["Symbol"]
            .str.strip()
            .str.replace(r"\.0+$", "", regex=True)
        )
        df_wz["Ilość_WZ"] = (
            df_wz["Ilość_WZ"].astype(str)
            .str.replace(",", ".", regex=False)
            .str.replace(r"\s+", "", regex=True)
        )
        df_wz["Ilość_WZ"] = pd.to_numeric(df_wz["Ilość_WZ"], errors="coerce").fillna(0)

    else:
        # ---------------------------------
        # 2a.2) Wariant B: jeśli nie ma kolumny "Ilość" wprost, to musimy złożyć ją z dwóch pól:
        #    - "Termin ważności Ilo" (część całkowita ilości po dacie)
        #    - "ść Waga brutto" (część dziesiętna + wagę brutto)
        # ---------------------------------
        col_part_int = next(
            (col for col in cols if "termin" in col.lower() and "ilo" in col.lower()),
            None
        )
        col_part_dec = next(
            (col for col in cols if "waga" in col.lower()),
            None
        )
        col_ean = next((col for col in cols if "kod" in col.lower() and "produkt" in col.lower()), None)

        if col_part_int is None or col_part_dec is None or col_ean is None:
            st.error(
                "Nie udało się wykryć kolumn w rozbitym układzie tabeli WZ (PDF).\n"
                "Spodziewane nagłówki: 'Kod produktu', 'Termin ważności Ilo', 'ść Waga brutto'.\n"
                f"Znalezione nagłówki: {cols}"
            )
            st.stop()

        # Teraz odtwórzmy 'Ilość_WZ' w każdym wierszu:
        eans = []
        ilosci = []
        for _, row in df_wz_raw.iterrows():
            raw_ean = str(row[col_ean]).strip()
            if raw_ean == "" or pd.isna(raw_ean):
                continue

            # Część całkowita: ostatni token w kolumnie col_part_int
            part_int_cell = str(row[col_part_int]).strip()
            tokens_int = part_int_cell.split()
            if len(tokens_int) < 2:
                int_part = "0"
            else:
                raw_int = tokens_int[-1]  # np. "150" lub "90"
                int_part = raw_int.replace(",", "").strip()

            # Część dziesiętna: pierwszy token w kolumnie col_part_dec
            part_dec_cell = str(row[col_part_dec]).strip()
            tokens_dec = part_dec_cell.split()
            if len(tokens_dec) == 0:
                dec_part = "00"
            else:
                dec_token = tokens_dec[0]  # np. ",00"
                dec_part = dec_token.replace(".", "").strip()  # usuwamy ewentualne kropki

            # Zbuduj string "150,00", zamień na "150.00" i skonwertuj
            if dec_part.startswith(","):
                qty_str = f"{int_part}{dec_part}"
            else:
                qty_str = f"{int_part},{dec_part}"
            qty_num = pd.to_numeric(qty_str.replace(",", "."), errors="coerce")
            if pd.isna(qty_num):
                qty_num = 0.0

            eans.append(raw_ean)
            ilosci.append(qty_num)

        df_wz = pd.DataFrame({
            "Symbol": eans,
            "Ilość_WZ": ilosci
        })

else:
    # ---------------------------------
    # 2b) Jeżeli wgrano gotowy Excel z WZ
    # ---------------------------------
    try:
        df_wz_raw = pd.read_excel(uploaded_wz, dtype={"Kod produktu": str})
    except Exception as e:
        st.error(f"Nie udało się wczytać pliku WZ (Excel):\n```{e}```")
        st.stop()

    if "Kod produktu" not in df_wz_raw.columns or "Ilość" not in df_wz_raw.columns:
        st.error(
            "Plik WZ (Excel) musi mieć kolumny:\n"
            "- `Kod produktu` (EAN)\n"
            "- `Ilość` (liczba sztuk w danym wierszu WZ)\n\n"
            f"Znalezione nagłówki: {list(df_wz_raw.columns)}"
        )
        st.stop()

    df_wz = df_wz_raw.rename(columns={"Kod produktu": "Symbol", "Ilość": "Ilość_WZ"})
    df_wz["Symbol"] = (
        df_wz["Symbol"].astype(str)
        .str.strip()
        .str.replace(r"\.0+$", "", regex=True)
    )
    df_wz["Ilość_WZ"] = (
        df_wz["Ilość_WZ"].astype(str)
        .str.replace(",", ".", regex=False)
        .str.replace(r"\s+", "", regex=True)
    )
    df_wz["Ilość_WZ"] = pd.to_numeric(df_wz["Ilość_WZ"], errors="coerce").fillna(0)

# =====================================
# 3) Grupowanie po EAN (Symbol) i sumowanie ilości
# =====================================
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

# =====================================
# 4) Scalanie (merge) i porównanie
# =====================================
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

status_order = ["Różni się", "Brak we WZ", "Brak w zamówieniu", "OK"]
df_compare["Status"] = pd.Categorical(
    df_compare["Status"], categories=status_order, ordered=True
)
df_compare = df_compare.sort_values(["Status", "Symbol"])

# =====================================
# 5) Wyświetlenie wyniku i przycisk do pobrania raportu
# =====================================
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
