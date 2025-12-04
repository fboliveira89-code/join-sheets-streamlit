import streamlit as st
from excel_join_like_macro_colA import join_sheets_like_macro_colA
import tempfile
import os

st.title("Join de Sheets em Excel")

uploaded = st.file_uploader("Carrega o ficheiro Excel", type=["xlsx"])

if uploaded:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(uploaded.read())
        tmp_path = tmp.name

    st.success("Ficheiro carregado com sucesso!")

    join_sheet = st.text_input("Nome da folha final (Join)", "Join")

    if st.button("Executar Join"):
        save_path = join_sheets_like_macro_colA(tmp_path, join_sheet)

        with open(save_path, "rb") as f:
            st.download_button(
                label="Download do Excel Resultado",
                data=f,
                file_name=f"Join_" + uploaded.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        st.success("Join concluído com sucesso!")
