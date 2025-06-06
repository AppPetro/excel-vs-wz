import pandas as pd

# 1) Ścieżki do plików (zmodyfikuj, jeżeli trzeba)
ORDER_XLSX = "MM do EPP.xlsx"           # plik z zamówieniem (kolumny: Symbol, Ilość)
WZ_XLSX   = "WZ-284_PET_HA_2025.xlsx"    # plik WZ po konwersji z PDF

# 2) Wczytanie zamówienia
# ------------------------------------------------------------------
df_order = pd.read_excel(ORDER_XLSX, dtype={"Symbol": str})
# Zakładamy, że kolumny w zamówieniu nazywają się dokładnie: "Symbol" i "Ilość"
# Jeżeli format jest inny, zmień nazwy poniżej (lub użyj rename).
df_order = df_order[["Symbol", "Ilość"]].copy()

# Usuń ewentualne spacje i sufiksy ".0" w EAN-ach
df_order["Symbol"] = (
    df_order["Symbol"]
    .str.strip()
    .str.replace(r"\.0+$", "", regex=True)
)

# 3) Wczytanie WZ (po konwersji PDF→Excel)
# ------------------------------------------------------------------
df_wz = pd.read_excel(WZ_XLSX, dtype={"Kod produktu": str})

# Przykładowe nazwy kolumn w WZ (po konwersji to może być "Kod produktu" i "Ilość"):
# Jeśli inne, to dostosuj poniżej.
df_wz = df_wz.rename(columns={"Kod produktu": "Symbol", "Ilość": "Ilość_WZ"})

# Usuń puste EAN-y, usuń sufiks ".0" i spacje, tak jak w zamówieniu
df_wz["Symbol"] = (
    df_wz["Symbol"]
    .str.strip()
    .str.replace(r"\.0+$", "", regex=True)
)

# 4) Grupowanie (zsumuj ilości dla tego samego EAN)
# ------------------------------------------------------------------
# Zamówienie – grupujemy, gdyby ktoś powtórzył EAN kilka razy w zamówieniu
df_order_grouped = (
    df_order
    .groupby("Symbol", as_index=False)
    .agg({"Ilość": "sum"})
    .rename(columns={"Ilość": "Zamówiona_ilość"})
)

# WZ – zsumuj wydaną ilość (zalążek: sumowanie ilości nawet gdy jest rozbite na różne terminy)
df_wz_grouped = (
    df_wz
    .groupby("Symbol", as_index=False)
    .agg({"Ilość_WZ": "sum"})
    .rename(columns={"Ilość_WZ": "Wydana_ilość"})
)

# 5) Połączenie tabel (merge) i porównanie
# ------------------------------------------------------------------
df_compare = pd.merge(
    df_order_grouped,
    df_wz_grouped,
    on="Symbol",
    how="outer",            # używamy outer, aby zobaczyć EAN-y tylko w zamówieniu i tylko w WZ
    indicator=True          # doda kolumnę "_merge" mówiącą, skąd wziął się każdy wiersz
)

# Wypełnij brakujące ilości 0 (jeśli jakiś EAN był tylko w zamówieniu, nie ma go we WZ – lub odwrotnie)
df_compare["Zamówiona_ilość"] = df_compare["Zamówiona_ilość"].fillna(0)
df_compare["Wydana_ilość"]    = df_compare["Wydana_ilość"].fillna(0)

# Dodaj kolumnę „Różnica” = Zamówiona - Wydana
df_compare["Różnica"] = df_compare["Zamówiona_ilość"] - df_compare["Wydana_ilość"]

# Dodaj kolumnę „Status”, czy się zgadza czy nie
def status_row(row):
    if row["Zamówiona_ilość"] == row["Wydana_ilość"] and row["_merge"] == "both":
        return "OK"
    elif row["_merge"] == "left_only":
        return "Brak we WZ"
    elif row["_merge"] == "right_only":
        return "Brak w zamówieniu"
    else:
        return "Różni się"

df_compare["Status"] = df_compare.apply(status_row, axis=1)

# 6) Posortuj po Symbolu lub po Statusie, jak wolisz
df_compare = df_compare.sort_values(["Status", "Symbol"], ascending=[True, True])

# 7) Zapisz wynik do Excela lub wypisz na ekran
# ------------------------------------------------------------------
print("\n===== Raport porównawczy Zamówienie vs WZ =====\n")
print(df_compare.to_string(index=False))

# Jeśli chcesz zapisać do pliku:
OUTPUT_XLSX = "porownanie_order_vs_wz.xlsx"
df_compare.to_excel(OUTPUT_XLSX, index=False)
print(f"\nZapisano raport do pliku: {OUTPUT_XLSX}")
