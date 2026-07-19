import streamlit as st
import sqlite3
import pandas as pd
import os

st.write("Banco:", os.path.abspath("usuarios_sistema.db"))

conn = sqlite3.connect("usuarios_sistema.db")
cursor = conn.cursor()

cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
st.write("Tabelas:", cursor.fetchall())

cursor.execute("SELECT username, nome, email, status FROM usuarios")
dados = cursor.fetchall()

st.write(f"Total de usuários: {len(dados)}")

if dados:
    st.dataframe(
        pd.DataFrame(
            dados,
            columns=["username", "nome", "email", "status"]
        )
    )

conn.close()
