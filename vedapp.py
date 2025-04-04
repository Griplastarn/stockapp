import streamlit as st
import math
from datetime import datetime
from openpyxl import Workbook, load_workbook
from fpdf import FPDF
import os
import base64
import pandas as pd

EXCEL_FIL = "vedlogg.xlsx"
PDF_FIL = "vedrapport.pdf"

def spara_till_excel(längd, diameter, volym, fast, travad):
    if os.path.exists(EXCEL_FIL):
        wb = load_workbook(EXCEL_FIL)
        ws = wb.active
        if ws["A" + str(ws.max_row)].value == "SUMMA":
            ws.delete_rows(ws.max_row)
    else:
        wb = Workbook()
        ws = wb.active
        ws.append(["Datum", "Längd (m)", "Diameter (cm)", "m³/stock", "m³fub", "m³s"])

    datum = datetime.now().strftime("%Y-%m-%d")
    ws.append([datum, längd, diameter, round(volym, 3), round(fast, 3), round(travad, 3)])

    total_fub = sum(row[4] for row in ws.iter_rows(min_row=2, values_only=True) if isinstance(row[4], (int, float)))
    total_s = sum(row[5] for row in ws.iter_rows(min_row=2, values_only=True) if isinstance(row[5], (int, float)))
    ws.append(["SUMMA", "", "", "", round(total_fub, 3), round(total_s, 3)])

    wb.save(EXCEL_FIL)

def skapa_pdf():
    try:
        wb = load_workbook(EXCEL_FIL)
        ws = wb.active

        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Helvetica", "B", 16)
        pdf.cell(0, 10, "Vedrapport", ln=True, align="C")

        pdf.set_font("Helvetica", "B", 12)
        pdf.ln(10)
        headers = ["Datum", "Längd", "Diameter", "m³/stock", "m³fub", "m³s"]
        col_widths = [50, 25, 25, 30, 30, 30]  # justerade bredder

        for i, h in enumerate(headers):
            pdf.cell(col_widths[i], 8, h, border=1)
        pdf.ln()

        pdf.set_font("Helvetica", "", 11)
        for row in ws.iter_rows(min_row=2, values_only=True):
            pdf.set_font("Helvetica", "B", 11) if row[0] == "SUMMA" else pdf.set_font("Helvetica", "", 11)
            for i, cell in enumerate(row[:6]):
                text = str(cell) if cell else ""
                pdf.cell(col_widths[i], 8, text, border=1)
            pdf.ln()

        pdf.output(PDF_FIL)
        return True
    except Exception:
        return False

def skapa_download_länk(filnamn, knapptext):
    with open(filnamn, "rb") as f:
        data = f.read()
    b64 = base64.b64encode(data).decode()
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="{filnamn}">{knapptext}</a>'
    return href

def rensa_data():
    if os.path.exists(EXCEL_FIL):
        os.remove(EXCEL_FIL)
    if os.path.exists(PDF_FIL):
        os.remove(PDF_FIL)
    st.success("All data har rensats!")

# --- Streamlit börjar här ---
st.set_page_config(page_title="Vedräknare", page_icon="🪵")
st.title("🪓 Vedräknare")

# Rensa-knapp
if st.button("🧹 Rensa allt"):
    rensa_data()

# Formulär
with st.form("vedform", clear_on_submit=True):
    längd_input = st.text_input("Längd på stock (meter)", value="", placeholder="Ex: 3.20")
    diameter_input = st.text_input("Diameter (cm)", value="", placeholder="Ex: 25.5")
    submitted = st.form_submit_button("Räkna och spara")

    if submitted:
        try:
            längd = float(längd_input.replace(",", "."))
            diameter = float(diameter_input.replace(",", "."))
            if längd > 0 and diameter > 0:
                radie = diameter / 200
                volym = math.pi * radie**2 * längd
                fast = volym
                travad = volym * 1.6

                spara_till_excel(längd, diameter, volym, fast, travad)
                st.success(f"✅ Volym: {volym:.3f} m³\nFast mått: {fast:.3f} m³fub\nTravad: {travad:.3f} m³s\nLoggat i vedlogg.xlsx")
            else:
                st.warning("❗ Värdena måste vara större än 0.")
        except ValueError:
            st.error("❌ Ange giltiga tal (punkt eller komma går bra).")

# Export till PDF
if st.button("📄 Exportera till PDF"):
    if skapa_pdf():
        st.success("📄 PDF skapad: vedrapport.pdf")
    else:
        st.error("❌ Fel vid PDF-export")

# Nedladdningslänkar
if os.path.exists(EXCEL_FIL):
    st.markdown(skapa_download_länk(EXCEL_FIL, "📥 Ladda ner Excel-fil"), unsafe_allow_html=True)
if os.path.exists(PDF_FIL):
    st.markdown(skapa_download_länk(PDF_FIL, "📥 Ladda ner PDF-rapport"), unsafe_allow_html=True)

# Visa datatabell
if os.path.exists(EXCEL_FIL):
    try:
        df = pd.read_excel(EXCEL_FIL)
        st.subheader("📊 Inmatade stockar")
        st.dataframe(df, use_container_width=True)
    except Exception as e:
        st.warning(f"Kunde inte läsa Excel-fil: {e}")
