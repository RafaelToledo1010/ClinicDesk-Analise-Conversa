"""
auth.py
Gerenciamento de usuários e autenticação
"""

import json
import hashlib
import os
from pathlib import Path

USERS_FILE = Path('users.json')

def hash_senha(senha):
    return hashlib.sha256(senha.encode()).hexdigest()

def carregar_usuarios():
    if not USERS_FILE.exists():
        # Usuários iniciais
        usuarios = {
            "rafael@clinicaclimes.com.br": {
                "senha": hash_senha("senha123"),
                "nome": "Rafael",
                "trocar_senha": True
            },
            "cliente@clinicdesk.com.br": {
                "senha": hash_senha("senha123"),
                "nome": "Cliente",
                "trocar_senha": True
            }
        }
        salvar_usuarios(usuarios)
    with open(USERS_FILE, encoding='utf-8') as f:
        return json.load(f)

def salvar_usuarios(usuarios):
    with open(USERS_FILE, 'w', encoding='utf-8') as f:
        json.dump(usuarios, f, ensure_ascii=False, indent=2)

def autenticar(email, senha):
    usuarios = carregar_usuarios()
    user = usuarios.get(email)
    if not user:
        return None
    if user['senha'] == hash_senha(senha):
        return user
    return None

def trocar_senha(email, senha_nova):
    usuarios = carregar_usuarios()
    if email not in usuarios:
        return False
    usuarios[email]['senha'] = hash_senha(senha_nova)
    usuarios[email]['trocar_senha'] = False
    salvar_usuarios(usuarios)
    return True
