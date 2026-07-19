import psycopg2
import os


def conectar_banco():

    conn = psycopg2.connect(
        os.environ["DATABASE_URL"]
    )

    cursor = conn.cursor()

    return conn, cursor
