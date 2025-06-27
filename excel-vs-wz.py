# -------------------------
# 2) Parsowanie zamówienia (Excel)
# -------------------------
try:
    # Wczytujemy bez nagłówka, bo sami go będziemy szukać
    df_order_raw = pd.read_excel(uploaded_order, dtype=str, header=None)
except Exception as e:
    st.error(f"Nie udało się wczytać pliku zamówienia:\n```{e}```")
    st.stop()

# Synonimy kolumn
syn_ean_ord = {
    normalize_col_name(c): c for c in [
        "Symbol","symbol","kod ean","ean","kod produktu","GTIN"
    ]
}
syn_qty_ord = {
    normalize_col_name(c): c for c in [
        "Ilość","Ilosc","Quantity","Qty","sztuki","ilość sztuk zamówiona","zamówiona ilość"
    ]
}

# Znajdź wiersz nagłówka
header_row = None
for idx, row in df_order_raw.iterrows():
    norm = [normalize_col_name(str(v)) for v in row.values]
    if any(h in syn_ean_ord for h in norm) and any(h in syn_qty_ord for h in norm):
        header = list(row.values)
        header_row = idx
        break

if header_row is None:
    st.error(
        "Excel Zlecenia/Zamówienia musi mieć w nagłówku kolumny EAN i Ilość.\n"
        f"Sprawdziłem wszystkie wiersze i nie znalazłem."
    )
    st.stop()

# Wyciągnij nazwy kolumn
col_ean_order = next(c for c in header if normalize_col_name(c) in syn_ean_ord)
col_qty_order = next(c for c in header if normalize_col_name(c) in syn_qty_ord)

# Parsuj dane poniżej nagłówka
ord_rows = []
for _, row in df_order_raw.iloc[header_row+1:].iterrows():
    raw_ean = str(row[col_ean_order]).strip().rstrip(".0")
    # oczyść ilość z białych znaków i zamień przecinek na kropkę
    raw_qty_str = str(row[col_qty_order]).strip()
    raw_qty_str = re.sub(r"\s+", "", raw_qty_str)  # usunięcie spacji, NBSP, tabów itp.
    raw_qty_str = raw_qty_str.replace(",", ".")
    # pomiń, jeśli nie ma w ogóle wartości ilości
    if raw_qty_str == "" or raw_qty_str.lower() in ("nan",):
        continue
    try:
        qty = float(raw_qty_str)
    except:
        # jeśli nie da się skonwertować, pomiń wiersz
        continue
    ord_rows.append([raw_ean, qty])

# Zbuduj df_order tylko z wierszy, które przeszły filtr
df_order = pd.DataFrame(ord_rows, columns=["Symbol", "Ilość"])
