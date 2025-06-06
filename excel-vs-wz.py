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
       - `Termin ważności Ilo` (zawiera datę i część całkowitą ilości)
       - `ść Waga brutto` (zawiera część po przecinku ilości i wagę)
    3. Aplikacja automatycznie:
       - wyciągnie tabelę z PDF-a za pomocą `pdfplumber`,
       - zidentyfikuje sposób zapisu ilości (normalny lub rozbity),
       - wyciągnie kolumny EAN (`Kod produktu`) i prawidłowo skonstruuje wartość `Ilość`,
       - zsumuje te ilości po EAN-ach,
       - porówna z zamówieniem i wygeneruje raport z różnicą.
    """
)

# Sidebar: upload plików
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
    - Jeśli wgrasz **PDF**, aplikacja użyje `pdfplumber` do wyciągnięcia tabeli i 
      rozpozna, czy kolumna „Ilość” jest od razu dostępna, czy rozbita na dwie części:
      * „Termin ważności Ilo” (część całkowita) i „ść Waga brutto” (część dziesiętna).  
    - Jeśli wgrasz **Excel** (plik już wyeksportowany ze WZ→.xlsx), 
      aplikacja odczyta kolumny `Kod produktu` i `Ilość` bezpośrednio.
    """
)

if uploaded_order is None or uploaded_wz is None:
    st.info("Proszę wgrać oba pliki po lewej stronie (zamówienie i WZ).")
    st.stop()

# -----------------------------------
# 1) Przetwarzanie zamówienia (Excel)
# -----------------------------------
try:
    df_order = pd.read_excel(uploaded_order, dtype={"Symbol": str})
except Exception as e:
    st.error(f"Nie udało się wczytać pliku zamówienia:\n```{e}```")
    st.stop()

if "Symbol" not in df_order.columns or "Ilość" not in df_order.columns:
    st.error(
        "Plik ZAMÓWIENIE musi mieć kolumny:\n"
        "- `Symbol` (EAN)\n"
        "- `Ilość` (liczba sztuk)\n\n"
        "Zweryfikuj, czy nagłówki dokładnie tak się nazywają (wielkość liter, spacje)."
    )
    st.stop()

# Oczyszczanie EAN-ów i konwersja ilości zamówionej na liczbę
df_order["Symbol"] = (
    df_order["Symbol"]
    .astype(str)
    .str.strip()
    .str.replace(r"\.0+$", "", regex=True)
)
df_order["Ilość"] = pd.to_numeric(df_order["Ilość"], errors="coerce").fillna(0)

# -----------------------------------
# 2) Przetwarzanie pliku WZ (PDF lub Excel)
# -----------------------------------
file_ext = uploaded_wz.name.lower().rsplit(".", maxsplit=1)[-1]

if file_ext == "pdf":
    # 2a) Ekstrakcja surowych tabel z PDF przy pomocy pdfplumber
    try:
        with pdfplumber.open(uploaded_wz) as pdf:
            all_tables = []
            for page in pdf.pages:
                extracted = page.extract_table()
                if extracted:
                    df_page = pd.DataFrame(extracted[1:], columns=extracted[0])
                    all_tables.append(df_page)
    except Exception as e:
        st.error(f"Nie udało się przeczytać PDF-a przez pdfplumber:\n```{e}```")
        st.stop()

    if len(all_tables) == 0:
        st.error("Nie znaleziono żadnych tabel w pliku PDF WZ.")
        st.stop()

    # Połącz wszystkie strony
    df_wz_raw = pd.concat(all_tables, ignore_index=True)

    # Sprawdź nagłówki w df_wz_raw.columns
    cols = list(df_wz_raw.columns)

    # Jeśli w nagłówkach jest bezpośrednio 'Ilość', użyjemy tej kolumny
    if any(col.lower().strip() == "ilość" or col.lower().strip() == "ilość " for col in cols):
        # Znajdź dokładną nazwę kolumny, która to 'Ilość'
        col_qty = next(col for col in cols if col.lower().strip().startswith("ilość"))
        col_ean = next((col for col in cols if "kod" in col.lower() and "produkt" in col.lower()), None)
        if col_ean is None:
            st.error(
                "Nie znaleziono kolumny 'Kod produktu' w pliku PDF WZ.\n"
                f"Znalezione nagłówki: {cols}"
            )
            st.stop()

        # Przygotuj DataFrame tylko z tych dwóch kolumn
        df_wz = pd.DataFrame({
            "Symbol": df_wz_raw[col_ean].astype(str),
            "Ilość_WZ": df_wz_raw[col_qty]
        })

        # Oczyść EAN i skonwertuj ilość na liczbę
        df_wz["Symbol"] = (
            df_wz["Symbol"]
            .str.strip()
            .str.replace(r"\.0+$", "", regex=True)
        )
        df_wz["Ilość_WZ"] = (
            df_wz["Ilość_WZ"]
            .astype(str)
            .str.replace(",", ".", regex=False)
            .str.replace(r"\s+", "", regex=True)
        )
        df_wz["Ilość_WZ"] = pd.to_numeric(df_wz["Ilość_WZ"], errors="coerce").fillna(0)

    else:
        # Zakładamy strukturę „rozbitą”:
        # Nagłówki: np. ['','Kod produktu','Nazwa','Termin ważności Ilo','ść Waga brutto']
        # Znajdź: kolumnę z 'Termin' i 'Ilo' (część całkowita), oraz kolumnę z 'Waga' (część dziesiętna)
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
                "Nie udało się dopasować rozbitej struktury kolumn w PDF WZ.\n"
                "Spodziewane kolumny: 'Kod produktu', 'Termin ważności Ilo', 'ść Waga brutto'.\n"
                f"Znalezione nagłówki: {cols}"
            )
            st.stop()

        # Teraz rekonstruujemy ilość w każdym wierszu:
        eans = []
        ilosci = []
        for _, row in df_wz_raw.iterrows():
            ean_raw = str(row[col_ean]).strip()
            if ean_raw == "" or pd.isna(ean_raw):
                continue
            # Część całkowita: ostatni token kolumny col_part_int (po dacie)
            part_int_cell = str(row[col_part_int])
            part_int_tokens = part_int_cell.strip().split()
            if len(part_int_tokens) < 2:
                # jeśli nie ma nic po dacie, załóż 0
                int_part = "0"
            else:
                raw_int = part_int_tokens[-1]
                int_part = raw_int.replace(",", "").strip()  # np. '150' lub '90,' → '150'/'90'
            # Część dziesiętna: pierwszy token kolumny col_part_dec (np. ',00 37,50' → ',00')
            part_dec_cell = str(row[col_part_dec])
            dec_token = part_dec_cell.strip().split()[0]  # np. ',00'
            dec_part = dec_token.replace(".", "").strip()  # nie powinno mieć kropki
            # Pełny string ilości, np. '150,00'
            qty_str = f"{int_part},{dec_part.lstrip(',')}" if dec_part.startswith(",") else f"{int_part},{dec_part}"
            # Zamiana na liczby (kropka = separator dziesiętny)
            qty_num = pd.to_numeric(qty_str.replace(",", "."), errors="coerce")
            if pd.isna(qty_num):
                qty_num = 0
            eans.append(ean_raw)
            ilosci.append(qty_num)

        df_wz = pd.DataFrame({
            "Symbol": eans,
            "Ilość_WZ": ilosci
        })

else:
    # 2b) Użytkownik wgrał gotowy Excel z WZ
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
            f"A zostały znalezione kolumny: {list(df_wz_raw.columns)}"
        )
        st.stop()

    # Zmień nazwy, oczyść i skonwertuj
    df_wz = df_wz_raw.rename(columns={"Kod produktu": "Symbol", "Ilość": "Ilość_WZ"})
    df_wz["Symbol"] = (
        df_wz["Symbol"]
        .astype(str)
        .str.strip()
        .str.replace(r"\.0+$", "", regex=True)
    )
    df_wz["Ilość_WZ"] = (
        df_wz["Ilość_WZ"]
        .astype(str)
        .str.replace(",", ".", regex=False)
        .str.replace(r"\s+", "", regex=True)
    )
    df_wz["Ilość_WZ"] = pd.to_numeric(df_wz["Ilość_WZ"], errors="coerce").fillna(0)

# -----------------------------------
# 3) Grupowanie po Symbol (EAN) – sumowanie ilości
# -----------------------------------
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

# -----------------------------------
# 4) Scalanie (merge) i obliczenie różnic
# -----------------------------------
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

# -----------------------------------
# 5) Wyświetlenie wyniku i pobranie raportu
# -----------------------------------
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
