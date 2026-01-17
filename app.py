import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import zipfile

st.set_page_config(page_title="Generator Dopisa", page_icon="📄")

st.title("📄 Generator Dopisa")
st.write("Automatizujte izradu Word dokumenata iz Excel tabele.")

# Polja za otpremanje
u_excel = st.file_uploader("1. Otpremi Excel tabelu (.xlsx)", type=["xlsx"])
u_word = st.file_uploader("2. Otpremi Word šablon (.docx)", type=["docx"])

if u_excel and u_word:
    df = pd.read_excel(u_excel)
    st.success("Podaci učitani!")
    st.dataframe(df.head(3)) # Prikazuje prva 3 reda
    
    # Korisnik bira kolonu po kojoj će se zvati fajlovi
    kolona_za_ime = st.selectbox("Izaberi kolonu za nazive fajlova:", df.columns)

    if st.button("Pokreni generisanje"):
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_f:
            for index, row in df.iterrows():
                doc = DocxTemplate(u_word)
                context = row.to_dict() # Pretvara red u podatke za Word
                doc.render(context)
                
                doc_io = io.BytesIO()
                doc.save(doc_io)
                
                naziv_fajla = f"{row[kolona_za_ime]}.docx"
                zip_f.writestr(naziv_fajla, doc_io.getvalue())
        
        st.download_button(
            label="⬇️ Preuzmi sve dopise (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="generisani_dopisi.zip"
        )
