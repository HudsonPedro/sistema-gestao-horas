import streamlit as st
from database import conectar_banco

st.title("Teste Neon")

try:
    conn, cursor = conectar_banco()

    cursor.execute("SELECT NOW();")
    resultado = cursor.fetchone()

    st.success("Conectado ao Neon!")
    st.write(resultado)

    conn.close()

except Exception as e:
    st.error(e)
