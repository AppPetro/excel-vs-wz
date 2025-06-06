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
    3. Aplikacja automatycznie:
       - wyciągnie tabelę z PDF-a za pomocą `pdfplumber`,
       - wyciągnie kolumny EAN (`Kod produktu`) i `Ilość`,
       - zsumuje te ilości po EAN-ach,
       - porówna z zamówieniem i wygeneruje raport z różnicą.
    """
)

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
    - Jeśli wgrasz **PDF**, aplikacja użyje `pdfplumber` do wyciągnięcia tabeli z kolumnami:
      `Kod produktu` (EAN) i `Ilość`.  
    - Jeśli wgrasz **Excel** (plik już wyeksportowany ręcznie ze WZ→.xlsx), 
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

# Oczyszczenie EAN-ów: usuń spacje, sufiks ".0"
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
    # 2a) Ekstrakcja tabeli z PDF za pomocą pdfplumber
    try:
        with pdfplumber.open(uploaded_wz) as pdf:
            all_tables = []
            for page in pdf.pages:
                # extract_table zwraca listę wierszy, gdzie wiersz to lista wartości
                table = page.extract_table()
                if table:
                    df_page = pd.DataFrame(table[1:], columns=table[0])
                    all_tables.append(df_page)
    except Exception as e:
        st.error(f"Nie udało się przeczytać PDF-a przez pdfplumber:\n```{e}```")
        st.stop()

    if len(all_tables) == 0:
        st.error("Nie znaleziono żadnych tabel w pliku PDF WZ.")
        st.stop()

    # Połącz wszystkie DataFrame-y w jeden
    df_wz_raw = pd.concat(all_tables, ignore_index=True)

    # Dopasuj kolumny: "Kod produktu" (EAN) i "Ilość"
    # Spróbuj wykryć nagłówek zawierający słowa "Kod" i "produkt"
    col_kod = None
    for col in df_wz_raw.columns:
        if "kod" in col.lower() and "produkt" in col.lower():
            col_kod = col
            break

    # Spróbuj wykryć kolumnę z "ilo" w nazwie (aby uchwycić "Ilość")
    col_ilosc = None
    for col in df_wz_raw.columns:
        if "ilo" in col.lower():
            col_ilosc = col
            break

    if col_kod is None or col_ilosc is None:
        st.error(
            "Po ekstrakcji PDF nie udało się odnaleźć kolumn:\n"
            "- `Kod produktu` (EAN)\n"
            "- `Ilość`\n\n"
            f"A zostały znalezione kolumny: {list(df_wz_raw.columns)}"
        )
        st.stop()

    df_wz = pd.DataFrame({
        "Symbol": df_wz_raw[col_kod].astype(str),
        "Ilość_WZ": df_wz_raw[col_ilosc]
    })

    # Oczyszczenie Symbol → usuń spacje, sufiks ".0"
    df_wz["Symbol"] = (
        df_wz["Symbol"]
        .str.strip()
        .str.replace(r"\.0+$", "", regex=True)
    )

    # Konwersja Ilość_WZ → float (jeśli były przecinki lub spacje w liczbie)
    df_wz["Ilość_WZ"] = (
        df_wz["Ilość_WZ"]
        .astype(str)
        .str.replace(",", ".", regex=False)
        .str.replace(r"\s+", "", regex=True)
        .astype(float, errors="coerce")
        .fillna(0)
    )

else:
    # 2b) Jeżeli użytkownik wgrał gotowy Excel z WZ
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
        .astype(float, errors="coerce")
        .fillna(0)
    )

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
# 4) Merge i porównanie
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

df_compare["Różnica"] = (
    df_compare["Zamówiona_ilość"] - df_compare["Wydana_ilość"]
)

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

# Posortuj (najpierw błędy, potem OK)
status_order = ["Różni się", "Brak we WZ", "Brak w zamówieniu", "OK"]
df_compare["Status"] = pd.Categorical(
    df_compare["Status"], categories=status_order, ordered=True
)
df_compare = df_compare.sort_values(["Status", "Symbol"])

# -----------------------------------
# 5) Wyświetlenie i pobranie wyniku
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
    writer.save()
    return output.getvalue()

st.download_button(
    label="⬇️ Pobierz raport jako Excel",
    data=to_excel(df_compare),
    file_name="porownanie_order_vs_wz.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.success("✅ Gotowe! Porównanie wykonane pomyślnie.")
