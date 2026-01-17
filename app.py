import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import zipfile

st.set_page_config(page_title="Generator Dopisa", page_icon="📄")

st.title("📄 Generator Dopisa iz Excela")
st.info("Ovaj alat menja markere poput {{ Naziv_kupca }} u Wordu podacima iz Excela.")

# Upload sekcija
u_excel = st.file_uploader("1. Otpremi Excel tabelu (.xlsx)", type=["xlsx"])
u_word = st.file_uploader("2. Otpremi Word šablon (.docx)", type=["docx"])

if u_excel and u_word:
    # Učitavanje podataka
    df = pd.read_excel(u_excel)
    
    # Čišćenje naziva kolona (menja razmak u donju crtu da bi Word lakše čitao)
    df.columns = [c.replace(' ', '_') for c in df.columns]
    
    st.success("Podaci su uspešno učitani!")
    st.write("Dostupne kolone (koristi ove nazive u Wordu unutar {{ }}):")
    st.code(", ".join(df.columns))
    
    # Izbor kolone za naziv fajla
    kolona_za_ime = st.selectbox("Izaberi kolonu za naziv generisanih fajlova:", df.columns)

    if st.button("🚀 Generiši dopise"):
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_f:
            for index, row in df.iterrows():
                doc = DocxTemplate(u_word)
                
                # Pretvaranje reda u rečnik (context)
                context = row.to_dict()
                
                # Popunjavanje šablona
                doc.render(context)
                
                # Čuvanje u memoriju
                doc_io = io.BytesIO()
                doc.save(doc_io)
                
                # Naziv fajla (uklanjanje nedozvoljenih karaktera)
                ime_fajla = str(row[kolona_za_ime]).replace("/", "-")
                zip_f.writestr(f"{ime_fajla}_{index}.docx", doc_io.getvalue())
        
        st.download_button(
            label="⬇️ Preuzmi ZIP sa dopisima",
            data=zip_buffer.getvalue(),
            file_name="generisani_dopisi.zip",
            mime="application/zip"
        )
