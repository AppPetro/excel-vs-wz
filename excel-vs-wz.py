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
    1. Wgraj plik Excel z zamówieniem, zawierający kolumny:
       - `Symbol` (EAN, 13 cyfr),
       - `Ilość` (liczba zamawianych sztuk).
    2. Wgraj plik WZ w formacie **PDF** (lub gotowy Excel z WZ), zawierający:
       - `Kod produktu` (EAN),
       - `Ilość` (wydane sztuki)
       LUB w PDF-ie:
       - rozbitą na dwie kolumny `Termin ważności Ilo` (część całkowita) i `ść Waga brutto` (część dziesiętna).
    3. Aplikacja:
       - wypakuje wszystkie strony PDF (ręcznie z pierwszej strony, przez regex, a z kolejnych przez `extract_tables`),  
       - odtworzy kolumnę `Ilość`,  
       - zsumuje po EAN-ach i porówna z zamówieniem.
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
    - Jeśli wgrasz **PDF**, aplikacja:
      • z pierwszej strony ręcznie wyciągnie EAN + ilość przez regex,  
      • z kolejnych stron użyje `extract_tables()`.  
    - Jeśli wgrasz **Excel** (WZ→.xlsx), wczyta `Kod produktu` + `Ilość` bezpośrednio.
    """
)

if uploaded_order is None or uploaded_wz is None:
    st.info("Proszę wgrać oba pliki: Excel z zamówieniem i PDF/Excel z WZ.")
    st.stop()

# =====================================
# 1) Wczytanie i przygotowanie Excela z ZAMÓWIENIEM
# =====================================
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
        "Sprawdź, czy nagłówki precyzyjnie się zgadzają."
    )
    st.stop()

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
    try:
        with pdfplumber.open(uploaded_wz) as pdf:
            all_tables = []

            def is_valid_wz_table(df: pd.DataFrame) -> bool:
                cols = [str(c).lower().strip() for c in df.columns]
                return any("kod" in c and "produkt" in c for c in cols) or any(c == "ilość" for c in cols)

            for page_idx, page in enumerate(pdf.pages):
                if page_idx == 0:
                    # ręczne wyciąganie z pierwszej strony:
                    text = page.extract_text() or ""
                    # znajdź wszystkie 13-cyfrowe EAN-y oraz liczby "xxx,xx"
                    eans = re.findall(r"\b(\d{13})\b", text)
                    qtys = re.findall(r"\b(\d{1,4},\d{2})\b", text)
                    # jeśli lista qtys krótsza od listy eans, dopasuj przez len=min
                    n = min(len(eans), len(qtys))
                    manual_rows = []
                    for i in range(n):
                        ean = eans[i]
                        qty_str = qtys[i].replace(",", ".")
                        try:
                            qty = float(qty_str)
                        except:
                            qty = 0.0
                        manual_rows.append([ean, qty])
                    if manual_rows:
                        df_manual = pd.DataFrame(manual_rows, columns=["Symbol", "Ilość_WZ"])
                        all_tables.append(df_manual)

                else:
                    # kolejne strony: extract_tables()
                    tables_on_page = page.extract_tables()
                    added = False
                    for table in tables_on_page:
                        if table and len(table) > 1:
                            df_page = pd.DataFrame(table[1:], columns=table[0])
                            if is_valid_wz_table(df_page):
                                all_tables.append(df_page)
                                added = True
                    if not added:
                        single = page.extract_table()
                        if single and len(single) > 1:
                            df_single = pd.DataFrame(single[1:], columns=single[0])
                            if is_valid_wz_table(df_single):
                                all_tables.append(df_single)
    except Exception as e:
        st.error(f"Nie udało się przeczytać PDF-a przez pdfplumber:\n```{e}```")
        st.stop()

    if len(all_tables) == 0:
        st.error("Nie znaleziono żadnej użytecznej tabeli w PDF WZ.")
        st.stop()

    df_wz_raw = pd.concat(all_tables, ignore_index=True)
    cols = list(df_wz_raw.columns)

    # Wariant A: bezpośrednia kolumna "Ilość"
    ilo_exists = next((col for col in cols if col.lower().strip() == "ilość"), None)
    if ilo_exists is not None:
        col_qty = ilo_exists
        col_ean = next((col for col in cols if "kod" in col.lower() and "produkt" in col.lower()), None)
        if col_ean is None:
            st.error(
                "Po połączeniu tabel nie znaleziono `Kod produktu`.\n"
                f"Znalezione nagłówki: {cols}"
            )
            st.stop()

        df_wz = pd.DataFrame({
            "Symbol": df_wz_raw[col_ean].astype(str).str.strip().apply(lambda x: x.split()[-1]),
            "Ilość_WZ": df_wz_raw[col_qty]
        })
        df_wz["Ilość_WZ"] = (
            df_wz["Ilość_WZ"].astype(str)
            .str.replace(",", ".", regex=False)
            .str.replace(r"\s+", "", regex=True)
        )
        df_wz["Ilość_WZ"] = pd.to_numeric(df_wz["Ilość_WZ"], errors="coerce").fillna(0)

    else:
        # Wariant B: rozbita kolumna ilości
        col_part_int = next((col for col in cols if "termin" in col.lower() and "ilo" in col.lower()), None)
        col_part_dec = next((col for col in cols if "waga" in col.lower()), None)
        col_ean = next((col for col in cols if "kod" in col.lower() and "produkt" in col.lower()), None)

        if col_part_int is None or col_part_dec is None or col_ean is None:
            st.error(
                "Brak wymaganych kolumn w rozbitym układzie.\n"
                f"Znalezione nagłówki: {cols}"
            )
            st.stop()

        eans = []
        ilosci = []
        for _, row in df_wz_raw.iterrows():
            raw_ean_cell = str(row[col_ean]).strip()
            if raw_ean_cell == "" or pd.isna(raw_ean_cell):
                continue
            raw_ean = raw_ean_cell.split()[-1]

            part_int_cell = str(row[col_part_int]).strip()
            tokens_int = part_int_cell.split()
            int_part = tokens_int[-1].replace(",", "").strip() if len(tokens_int) >= 2 else "0"

            part_dec_cell = str(row[col_part_dec]).strip()
            tokens_dec = part_dec_cell.split()
            dec_part = tokens_dec[0].replace(".", "").strip() if tokens_dec else "00"

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
    # wczytanie gotowego Excela z WZ
    try:
        df_wz_raw = pd.read_excel(uploaded_wz, dtype={"Kod produktu": str})
    except Exception as e:
        st.error(f"Nie udało się wczytać pliku WZ (Excel):\n```{e}```")
        st.stop()

    if "Kod produktu" not in df_wz_raw.columns or "Ilość" not in df_wz_raw.columns:
        st.error(
            "Plik WZ (Excel) musi mieć kolumny:\n"
            "- `Kod produktu` (EAN)\n"
            "- `Ilość` (liczba sztuk)\n\n"
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
# 4) Scalanie (merge) i obliczenie różnic
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
