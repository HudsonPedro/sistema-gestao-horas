import streamlit as st
import sqlite3
import hashlib
import os

st.title("Reset da senha do Admin")

st.write("Banco:", os.path.abspath("usuarios_sistema.db"))

def criptografar_senha(senha):
    return hashlib.sha256(senha.encode()).hexdigest()

if st.button("Redefinir senha do admin para Admin@123"):

    try:
        conn = sqlite3.connect("usuarios_sistema.db")
        cursor = conn.cursor()

        cursor.execute(
            "UPDATE usuarios SET senha_hash=? WHERE username='admin'",
            (criptografar_senha("Admin@123"),)
        )

        conn.commit()

        st.success(f"Linhas alteradas: {cursor.rowcount}")

        conn.close()

    except Exception as e:
        st.error(e)
