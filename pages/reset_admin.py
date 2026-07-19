import streamlit as st
import sqlite3
import hashlib

def criptografar_senha(senha):
    return hashlib.sha256(senha.encode()).hexdigest()

if st.button("Criar administrador"):

    conn = sqlite3.connect("usuarios_sistema.db")
    cursor = conn.cursor()

    cursor.execute("""
        INSERT INTO usuarios
        (username,nome,senha_hash,email,status)
        VALUES (?,?,?,?,?)
    """, (
        "admin",
        "Administrador",
        criptografar_senha("Admin@123"),
        "admin@hptech.com",
        "Ativo"
    ))

    conn.commit()
    conn.close()

    st.success("Administrador criado!")
