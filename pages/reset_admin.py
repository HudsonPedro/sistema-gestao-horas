import sqlite3
import hashlib

conn = sqlite3.connect("usuarios_sistema.db")
cursor = conn.cursor()

cursor.execute("""
INSERT INTO usuarios
(username, nome, senha_hash, email, status)
VALUES (?,?,?,?,?)
""", (
    "admin",
    "Administrador",
    hashlib.sha256("Admin@123".encode()).hexdigest(),
    "hptech@hptechinformatica.com",
    "Ativo"
))

conn.commit()
conn.close()
