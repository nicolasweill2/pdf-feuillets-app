import streamlit as st
import zipfile
import os
import tempfile
from process_pdfs import process_folder  # ta fonction principale qui traite les PDFs

st.title("Analyseur de PDF d'impression (Feuillets)")

uploaded_zip = st.file_uploader("Chargez un dossier .zip contenant vos PDF", type="zip")

if uploaded_zip:
    with tempfile.TemporaryDirectory() as tmpdir:
        zip_path = os.path.join(tmpdir, "archive.zip")
        with open(zip_path, "wb") as f:
            f.write(uploaded_zip.read())

        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(tmpdir)

        pdf_folder = tmpdir  # Dossier temporaire où les PDFs sont extraits

        st.write("Traitement en cours...")
        output_excel_path = process_folder(pdf_folder)  # ta fonction renvoyant le chemin du fichier Excel
        st.success("Traitement terminé !")

        with open(output_excel_path, "rb") as f:
            st.download_button(
                label="📥 Télécharger le fichier Excel généré",
                data=f,
                file_name=os.path.basename(output_excel_path),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
