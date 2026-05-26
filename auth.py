"""
auth.py
Gerenciamento de autenticação — senhas fixas no código
"""

import hashlib

def hash_senha(senha):
    return hashlib.sha256(senha.encode()).hexdigest()

# Usuários cadastrados — edite aqui para adicionar/alterar senhas
USUARIOS = {
    "rafaeltoledo@clinicdesk.com.br": {
        "senha_hash": hash_senha("clinicdesk2026"),
        "nome": "Rafael"
    },
    "cliente@clinicdesk.com.br": {
        "senha_hash": hash_senha("clinicdesk2026"),
        "nome": "Cliente"
    }
}

def autenticar(email, senha):
    user = USUARIOS.get(email)
    if not user:
        return None
    if user["senha_hash"] == hash_senha(senha):
        return user
    return None

def trocar_senha(email, senha_nova):
    # No arquivo fixo não persiste — instrua o admin a editar o código
    return True
