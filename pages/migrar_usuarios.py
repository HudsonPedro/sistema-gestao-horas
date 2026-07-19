import sqlite3
import streamlit as st
import os

st.title("Diagnóstico SQLite")

st.write("Arquivo encontrado:")
st.write(os.path.abspath("usuarios_sistema.db"))

conn = sqlite3.connect("usuarios_sistema.db")
cursor = conn.cursor()

cursor.execute("""
SELECT name 
FROM sqlite_master 
WHERE type='table'
""")

st.write("Tabelas encontradas:")
st.write(cursor.fetchall())

conn.close()


# ==============================
# SQLite antigo
# ==============================

sqlite_conn = sqlite3.connect(
    "usuarios_sistema.db"
)

sqlite_cursor = sqlite_conn.cursor()


# ==============================
# Neon PostgreSQL
# ==============================

postgres_conn = psycopg2.connect(
    "postgresql://neondb_owner:npg_XdU6cRYoJpi9@ep-restless-term-au36ashx-pooler.c-10.us-east-1.aws.neon.tech/neondb?sslmode=require"
)

postgres_cursor = postgres_conn.cursor()


# ==============================
# Buscar usuários
# ==============================

sqlite_cursor.execute("""
SELECT 
    username,
    nome,
    senha_hash,
    email,
    status
FROM usuarios
""")

usuarios = sqlite_cursor.fetchall()


print("Usuários encontrados:", len(usuarios))


# ==============================
# Inserir no Neon
# ==============================

for usuario in usuarios:

    postgres_cursor.execute("""
    INSERT INTO usuarios
    (
        username,
        nome,
        senha_hash,
        email,
        status
    )
    VALUES (%s,%s,%s,%s,%s)

    ON CONFLICT(username)
    DO NOTHING

    """, usuario)


postgres_conn.commit()


sqlite_conn.close()
postgres_conn.close()


print("Migração concluída!")
