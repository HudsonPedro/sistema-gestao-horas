import streamlit as st
import hashlib

st.title("Gerar Hash")

senha = st.text_input(
    "Digite a senha do admin",
    type="password"
)

if st.button("Gerar"):
    hash_senha = hashlib.sha256(
        senha.encode()
    ).hexdigest()

    st.write(hash_senha)
