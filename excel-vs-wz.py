import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="Porównywarka Zamówienie ↔ WZ",
    layout="wide",
)

st.title("📋 Porównywarka Zamówienie (Excel) vs. WZ (Excel)")

st.markdown(
    """
    Ten prosty formularz pozwala wgrać:
    1. Excel z zamówieniem (kolumny: `Symbol` [EAN] i `Ilość`),
    2. Excel z WZ (po konwersji z PDF, kolumny: `Kod produktu` [EAN] i `Ilość`).
    
    Aplikacja:
    - zsumuje ilość w obu plikach po kolumnie EAN,
    - porówna, ile sztuk zamówiono, a ile faktycznie wydano,
    - i wygeneruje tabelę z różnicami i statusem każdej pozycji.
    """
)

# -----------------------------------
# 1) FileUploader: wczytanie Excela
# -----------------------------------
st.sidebar.header("Wgraj pliki Excel")

uploaded_order = st.sidebar.file_uploader(
    label="Wybierz plik ZAMÓWIENIE (xlsx)",
    type=["xlsx"],
    key="order_uploader"
)

uploaded_wz = st.sidebar.file_uploader(
    label="Wybierz plik WZ (xlsx)",
    type=["xlsx"],
    key="wz_uploader"
)

st.sidebar.markdown(
    """
    **Instrukcja:**
    - Najpierw przygotuj plik Excel z zamówieniem (np. „MM do EPP.xlsx”), 
      w którym są przynajmniej kolumny:
      - `Symbol` (EAN, np. 5029040012281)
      - `Ilość` (liczba zamawianych sztuk)
    - Następnie przygotuj plik Excel z WZ (po konwersji z PDF),
      w którym są przynajmniej kolumny:
      - `Kod produktu` (EAN)
      - `Ilość` (liczba wydanych sztuk w danym wierszu WZ)
    - Po wgraniu obu plików, aplikacja wyprodukuje porównanie.
    """
)

# -----------------------------------
# 2) Gdy oba pliki wgrane – wykonaj analizę
# -----------------------------------
if uploaded_order is not None and uploaded_wz is not None:
    try:
        # 2a) Wczytanie zamówienia
        df_order = pd.read_excel(uploaded_order, dtype={"Symbol": str})
        # Upewniamy się, że mamy kolumny z odpowiednimi nazwami:
        if "Symbol" not in df_order.columns or "Ilość" not in df_order.columns:
            st.error(
                "Plik zamówienia musi mieć kolumny: "
                "`Symbol` (EAN) oraz `Ilość` (ilość sztuk)."
            )
            st.stop()

        # Oczyszczenie EAN-ów (usuń spacje i końcówki „.0”)
        df_order["Symbol"] = (
            df_order["Symbol"]
            .astype(str)
            .str.strip()
            .str.replace(r"\.0+$", "", regex=True)
        )
        df_order["Ilość"] = pd.to_numeric(df_order["Ilość"], errors="coerce").fillna(0)

        # 2b) Wczytanie WZ
        df_wz = pd.read_excel(uploaded_wz, dtype={"Kod produktu": str})
        # Sprawdź nazwy kolumn:
        if "Kod produktu" not in df_wz.columns or "Ilość" not in df_wz.columns:
            st.error(
                "Plik WZ musi mieć kolumny: "
                "`Kod produktu` (EAN) oraz `Ilość` (ilość sztuk w danym wierszu WZ)."
            )
            st.stop()

        # Zmieniamy nazwy kolumn w WZ, aby pasowały do df_order
        df_wz = df_wz.rename(columns={"Kod produktu": "Symbol", "Ilość": "Ilość_WZ"})
        df_wz["Symbol"] = (
            df_wz["Symbol"]
            .astype(str)
            .str.strip()
            .str.replace(r"\.0+$", "", regex=True)
        )
        # Jeśli w kolumnie Ilość_WZ są przecinki („150,00”), zamień na kropkę i na float:
        df_wz["Ilość_WZ"] = (
            df_wz["Ilość_WZ"]
            .astype(str)
            .str.replace(",", ".", regex=False)
            .astype(float)
            .fillna(0)
        )

        # -----------------------------------
        # 3) Grupowanie po EAN-ach (sumowanie ilości)
        # -----------------------------------
        # Dla zamówienia:
        df_order_grouped = (
            df_order
            .groupby("Symbol", as_index=False)
            .agg({"Ilość": "sum"})
            .rename(columns={"Ilość": "Zamówiona_ilość"})
        )

        # Dla WZ:
        df_wz_grouped = (
            df_wz
            .groupby("Symbol", as_index=False)
            .agg({"Ilość_WZ": "sum"})
            .rename(columns={"Ilość_WZ": "Wydana_ilość"})
        )

        # -----------------------------------
        # 4) Połączenie (merge) i obliczenie różnic
        # -----------------------------------
        df_compare = pd.merge(
            df_order_grouped,
            df_wz_grouped,
            on="Symbol",
            how="outer",
            indicator=True
        )

        # Uzupełnij 0 tam, gdzie brakuje wartości
        df_compare["Zamówiona_ilość"] = df_compare["Zamówiona_ilość"].fillna(0)
        df_compare["Wydana_ilość"]    = df_compare["Wydana_ilość"].fillna(0)

        # Oblicz różnicę
        df_compare["Różnica"] = (
            df_compare["Zamówiona_ilość"] - df_compare["Wydana_ilość"]
        )

        # Dodaj kolumnę Status
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

        # Posortuj (opcjonalnie najpierw błędy, potem OK)
        status_order = ["Różni się", "Brak we WZ", "Brak w zamówieniu", "OK"]
        df_compare["Status"] = pd.Categorical(
            df_compare["Status"], categories=status_order, ordered=True
        )
        df_compare = df_compare.sort_values(["Status", "Symbol"])

        # -----------------------------------
        # 5) Wyświetlenie wyniku w Streamlit
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

        # 5a) Daj możliwość pobrania raportu jako Excel
        def to_excel(df: pd.DataFrame) -> bytes:
            """Zwraca bajty pliku .xlsx z pandas DataFrame."""
            from io import BytesIO
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

    except Exception as e:
        st.error(f"Wystąpił błąd podczas przetwarzania danych:\n```{e}```")
        st.stop()
else:
    st.info("Proszę wgrać oba pliki Excel po lewej stronie (zamówienie i WZ).")
