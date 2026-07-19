import sqlite3
import psycopg2


# SQLite atual
sqlite = sqlite3.connect("usuarios_sistema.db")
sqlite_cursor = sqlite.cursor()


# Neon PostgreSQL
postgres = psycopg2.connect(
    "postgresql://neondb_owner:npg_XdU6cRYoJpi9@ep-restless-term-au36ashx-pooler.c-10.us-east-1.aws.neon.tech/neondb?sslmode=require"
)

pg_cursor = postgres.cursor()


# Buscar usuários antigos
sqlite_cursor.execute("""
SELECT username, nome, senha_hash, email, status
FROM usuarios
""")

usuarios = sqlite_cursor.fetchall()


for usuario in usuarios:

    pg_cursor.execute("""
    INSERT INTO usuarios
    (username, nome, senha_hash, email, status)
    VALUES (%s,%s,%s,%s,%s)
    ON CONFLICT (username)
    DO NOTHING
    """, usuario)


postgres.commit()

sqlite.close()
postgres.close()


print(
    f"{len(usuarios)} usuários migrados com sucesso."
)
