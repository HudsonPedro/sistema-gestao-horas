import streamlit as st
import sqlite3
import hashlib

def criptografar_senha(senha):
    return hashlib.sha256(senha.encode()).hexdigest()

conn = sqlite3.connect("usuarios_sistema.db")
cursor = conn.cursor()

nova_senha = "Admin@123"

cursor.execute(
    "UPDATE usuarios SET senha_hash=? WHERE username='admin'",
    (criptografar_senha(nova_senha),)
)

conn.commit()
conn.close()

print("Senha do admin alterada com sucesso.")
