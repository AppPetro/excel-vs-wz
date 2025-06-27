
**Wynik:**  
- Tabela: **Symbol**, **Zamówiona_ilość**, **Wydana_ilość**, **Różnica**, **Status**.  
- Zielone wiersze = OK; czerwone = rozbieżności/braki.  
- Kliknij „⬇️ Pobierz raport”, by pobrać gotowy plik Excel.
      """
  )

st.sidebar.header("Krok 1: Zlecenie/Zamówienie")
up1 = st.sidebar.file_uploader("Wybierz plik", type=["xlsx","pdf"], key="file1")
st.sidebar.header("Krok 2: WZ")
up2 = st.sidebar.file_uploader("Wybierz plik", type=["xlsx","pdf"], key="file2")

if not up1 or not up2:
  st.info("Proszę wgrać oba pliki.")
  st.stop()

# ── Synonimy kolumn ─────────────────────────────────────────────
EAN_SYNS = ["Symbol","symbol","kod ean","ean","kod produktu","gtin"]
QTY_SYNS = ["Ilość","Ilosc","Quantity","Qty","sztuki","ilość sztuk zamówiona","zamówiona ilość"]

# ── Parsowanie pierwszego pliku ─────────────────────────────────
if up1.name.lower().endswith(".xlsx"):
  df1 = parse_excel(up1, EAN_SYNS, QTY_SYNS, "Ilość_Zam")
else:
  df1 = parse_pdf(up1, "Ilość_Zam")

# ── Parsowanie drugiego pliku ──────────────────────────────────
if up2.name.lower().endswith(".xlsx"):
  df2 = parse_excel(up2, EAN_SYNS, QTY_SYNS, "Ilość_WZ")
else:
  df2 = parse_pdf(up2, "Ilość_WZ")

# ── Grupowanie i porównanie ────────────────────────────────────
g1 = df1.groupby("Symbol", as_index=False).sum().rename(columns={"Ilość_Zam":"Zamówiona_ilość"})
g2 = df2.groupby("Symbol", as_index=False).sum().rename(columns={"Ilość_WZ":"Wydana_ilość"})
cmp = pd.merge(g1, g2, on="Symbol", how="outer", indicator=True)
cmp["Zamówiona_ilość"] = cmp["Zamówiona_ilość"].fillna(0)
cmp["Wydana_ilość"]    = cmp["Wydana_ilość"].fillna(0)
cmp["Różnica"]         = cmp["Zamówiona_ilość"] - cmp["Wydana_ilość"]

def status(r):
  if r["_merge"] == "left_only":
      return "Brak we WZ"
  if r["_merge"] == "right_only":
      return "Brak w zamówieniu"
  return "OK" if r["Różnica"] == 0 else "Różni się"

cmp["Status"] = cmp.apply(status, axis=1)
order = ["Różni się","Brak we WZ","Brak w zamówieniu","OK"]
cmp["Status"] = pd.Categorical(cmp["Status"], categories=order, ordered=True)
cmp.sort_values(["Status","Symbol"], inplace=True)

def highlight_row(row):
  color = "#c6efce" if row["Status"] == "OK" else "#ffc7ce"
  return [f"background-color: {color}"] * len(row)

st.markdown("### 📊 Wynik porównania")
st.dataframe(
  cmp.style
     .format({"Zamówiona_ilość":"{:.0f}", "Wydana_ilość":"{:.0f}", "Różnica":"{:.0f}"})
     .apply(highlight_row, axis=1),
  use_container_width=True
)

buf = BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as writer:
  cmp.to_excel(writer, index=False, sheet_name="Porównanie")

st.download_button("⬇️ Pobierz raport",
  data=buf.getvalue(),
  file_name="raport.xlsx",
  mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

if (cmp["Status"] == "OK").all():
  st.markdown("<h4 style='color:green;'>✅ Pozycje się zgadzają</h4>", unsafe_allow_html=True)
else:
  st.markdown("<h4 style='color:red;'>❌ Pozycje się nie zgadzają</h4>", unsafe_allow_html=True)
