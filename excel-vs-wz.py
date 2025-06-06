import pdfplumber
import pandas as pd

# ... wcześniej w app.py masz wczytanie uploaded_wz ...

file_ext = uploaded_wz.name.lower().rsplit(".", maxsplit=1)[-1]

if file_ext == "pdf":
    # Poprzednio używaliśmy extract_table() – ono czasem pobiera tylko pierwszą
    # wykrytą tabelę ze strony. Teraz zastępujemy to extract_tables(), 
    # żeby wyciągnąć wszystkie tabele ze wszystkich stron.
    try:
        with pdfplumber.open(uploaded_wz) as pdf:
            all_tables = []
            for page in pdf.pages:
                # extract_tables() zwraca listę wszystkich tabel znalezionych na stronie
                tables_on_page = page.extract_tables()
                for table in tables_on_page:
                    # Jeśli tabela nie jest pusta, przerabiamy ją na DataFrame
                    if table and len(table) > 1:
                        df_page = pd.DataFrame(table[1:], columns=table[0])
                        all_tables.append(df_page)
    except Exception as e:
        st.error(f"Nie udało się przeczytać PDF-a przez pdfplumber:\n```{e}```")
        st.stop()

    if len(all_tables) == 0:
        st.error("Nie znaleziono żadnych tabel w pliku PDF WZ.")
        st.stop()

    # Połącz wszystkie DataFrame’y w jeden – dzięki temu uzyskamy TYLKO JEDNĄ wspólną tabelę,
    # zawierającą wiersze ze wszystkich stron PDF.
    df_wz_raw = pd.concat(all_tables, ignore_index=True)

    # Teraz w df_wz_raw mamy wszystkie wiersze z każdej strony. 
    # Możemy dalej wykrywać, czy użyć kolumny "Ilość" czy tej rozbitych:
    cols = list(df_wz_raw.columns)

    # 1) Jeśli znajdziemy nagłówek "Ilość" – to jest zwykły layout (np. strona 2 i 3 WZ-284…)
    if any(col.lower().strip() == "ilość" for col in cols):
        col_qty    = next(col for col in cols if col.lower().strip() == "ilość")
        col_ean    = next((col for col in cols if "kod" in col.lower() and "produkt" in col.lower()), None)
        if col_ean is None:
            st.error(
                "Po scaleniu tabel nie znaleziono kolumny 'Kod produktu' w PDF WZ.\n"
                f"Znalezione nagłówki: {cols}"
            )
            st.stop()

        df_wz = pd.DataFrame({
            "Symbol": df_wz_raw[col_ean].astype(str),
            "Ilość_WZ": df_wz_raw[col_qty]
        })

        # Oczyść EAN i skonwertuj "Ilość_WZ" na liczbę
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
        # 2) Jeżeli nie ma bezpośredniej kolumny "Ilość", to musimy złożyć ją z dwóch pól:
        #    - kolumna zawierająca "Termin ważności Ilo" (część całkowita ilości po dacie)
        #    - kolumna zawierająca "ść Waga brutto" (część po przecinku + wagę brutto)
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
                "Nie udało się dopasować rozbitej struktury w PDF WZ.\n"
                "Spodziewane kolumny: 'Kod produktu', 'Termin ważności Ilo', 'ść Waga brutto'.\n"
                f"Znalezione nagłówki: {cols}"
            )
            st.stop()

        # Rekonstrukcja „Ilość” z dwóch kolumn:
        eans = []
        ilosci = []
        for _, row in df_wz_raw.iterrows():
            ean_raw = str(row[col_ean]).strip()
            if ean_raw == "" or pd.isna(ean_raw):
                continue

            # 2a) Część całkowita – ostatni token w kolumnie col_part_int, np. "2027-11-27 150"
            part_int_cell = str(row[col_part_int]).strip()
            part_int_tokens = part_int_cell.split()
            if len(part_int_tokens) < 2:
                int_part = "0"
            else:
                raw_int = part_int_tokens[-1]
                int_part = raw_int.replace(",", "").strip()

            # 2b) Część dziesiętna – pierwszy token kolumny col_part_dec, np. ",00 37,50"
            part_dec_cell = str(row[col_part_dec]).strip()
            dec_token = part_dec_cell.split()[0]  # np. ",00"
            dec_part = dec_token.replace(".", "").strip()  # np. "00" lub ",00"

            # Zbuduj string typu "150,00" i skonwertuj na float → 150.00
            # (zastępujemy przecinek na kropkę)
            if dec_part.startswith(","):
                qty_str = f"{int_part}{dec_part}"
            else:
                qty_str = f"{int_part},{dec_part}"
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
    # Jeżeli użytkownik wgrał gotowy Excel z WZ
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
