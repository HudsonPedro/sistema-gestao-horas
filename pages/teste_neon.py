import streamlit as st
import psycopg2

st.title("Teste conexão Neon")

try:
    conn = psycopg2.connect(
        st.secrets["neon"]["DATABASE_URL"]
    )

    cursor = conn.cursor()

    cursor.execute("SELECT NOW();")
    resultado = cursor.fetchone()

    st.success("Conectado ao Neon com sucesso!")
    st.write("Data do servidor:")
    st.write(resultado)

    conn.close()

except Exception as e:
    st.error("Erro na conexão:")
    st.write(e)
