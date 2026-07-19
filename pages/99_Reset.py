import streamlit as st
import sqlite3
import hashlib

st.title("🔄 Recuperação de Emergência HPTECH")

# Executa a alteração direto no arquivo físico do banco de dados no servidor
try:
    conn = sqlite3.connect("usuarios_sistema.db")
    cursor = conn.cursor()
    
    # Hash SHA-256 exato para a senha: Den559hurt301*
    nova_senha_hash = hashlib.sha256("Den559hurt301*".encode()).hexdigest()
    
    # Força a atualização do admin no arquivo .db do servidor
    cursor.execute("""
        UPDATE usuarios 
        SET senha_hash = ?, status = 'Ativo' 
        WHERE username = 'admin'
    """, (nova_senha_hash,))
    
    conn.commit()
    conn.close()
    st.success("✅ Sucesso! A senha do usuário 'admin' foi alterada para: Den559hurt301*")
    st.info("Agora você já pode voltar para a tela inicial e fazer o login.")
except Exception as e:
    st.error(f"Erro ao acessar o banco: {e}")
