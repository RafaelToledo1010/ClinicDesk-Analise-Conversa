"""
app.py
Servidor web Flask — Clínica Climes
Variáveis de ambiente: ANTHROPIC_API_KEY, SECRET_KEY
Endpoints expostos: /, /login, /logout, /trocar-senha, /processar, /analisar, /download, /status
"""

import json
import os
import threading
import uuid
from datetime import datetime
from functools import wraps
from pathlib import Path

from flask import (
    Flask, render_template, request,
    redirect, url_for, send_file, jsonify, session
)

from estruturar import estruturar
from analisar import analisar_todas, gerar_excel
from auth import autenticar, trocar_senha

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'clinica-climes-2026')

ANTHROPIC_API_KEY = os.environ.get('ANTHROPIC_API_KEY', '')

UPLOAD_FOLDER = Path('uploads')
OUTPUT_FOLDER = Path('outputs')
UPLOAD_FOLDER.mkdir(exist_ok=True)
OUTPUT_FOLDER.mkdir(exist_ok=True)

jobs = {}


# ─────────────────────────────────────────────
# AUTENTICAÇÃO
# ─────────────────────────────────────────────

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'usuario' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = request.form.get('email', '').strip().lower()
        senha = request.form.get('senha', '')
        user = autenticar(email, senha)
        if not user:
            return render_template('login.html', erro='E-mail ou senha incorretos.')
        session['usuario'] = email
        session['nome'] = user.get('nome', email)
        if user.get('trocar_senha'):
            return redirect(url_for('trocar_senha_route'))
        return redirect(url_for('index'))
    return render_template('login.html')


@app.route('/trocar-senha', methods=['GET', 'POST'])
def trocar_senha_route():
    if 'usuario' not in session:
        return redirect(url_for('login'))
    if request.method == 'POST':
        senha_nova = request.form.get('senha_nova', '')
        senha_confirma = request.form.get('senha_confirma', '')
        if len(senha_nova) < 6:
            return render_template('trocar_senha.html', erro='A senha deve ter pelo menos 6 caracteres.')
        if senha_nova != senha_confirma:
            return render_template('trocar_senha.html', erro='As senhas não coincidem.')
        trocar_senha(session['usuario'], senha_nova)
        return redirect(url_for('index'))
    return render_template('trocar_senha.html')


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


# ─────────────────────────────────────────────
# ROTAS PRINCIPAIS
# ─────────────────────────────────────────────

@app.route('/')
@login_required
def index():
    historico = []
    for f in sorted(OUTPUT_FOLDER.glob('*.json'), reverse=True)[:10]:
        try:
            with open(f, encoding='utf-8') as fp:
                dados = json.load(fp)
            historico.append({
                'nome': f.stem,
                'data': datetime.fromtimestamp(f.stat().st_mtime).strftime('%d/%m/%Y %H:%M'),
                'total': len(dados),
                'agendaram': sum(1 for c in dados if c.get('gerou_agendamento') is True),
                'nao_agendaram': sum(1 for c in dados if c.get('gerou_agendamento') is False),
            })
        except Exception:
            continue
    return render_template('index.html', historico=historico, nome=session.get('nome',''))


@app.route('/processar', methods=['GET', 'POST'])
@login_required
def processar():
    if request.method == 'POST':
        arquivo = request.files.get('arquivo')
        if not arquivo or not arquivo.filename.endswith(('.xlsx', '.xls')):
            return render_template('processar.html', erro='Selecione um arquivo Excel válido.')
        nome = Path(arquivo.filename).stem
        caminho = UPLOAD_FOLDER / arquivo.filename
        arquivo.save(str(caminho))
        try:
            def log_noop(msg): pass
            conversas, stats = estruturar(str(caminho), log_noop)
            saida_json = OUTPUT_FOLDER / f'{nome}.json'
            with open(saida_json, 'w', encoding='utf-8') as f:
                json.dump(conversas, f, ensure_ascii=False, indent=2, default=str)
            session['arquivo_atual'] = nome
            session['stats'] = stats
            return render_template('processar.html', sucesso=True, nome=nome, stats=stats)
        except Exception as e:
            return render_template('processar.html', erro=f'Erro ao processar: {str(e)}')
    return render_template('processar.html')


@app.route('/analisar/<nome>', methods=['POST'])
@login_required
def analisar(nome):
    if not ANTHROPIC_API_KEY:
        return jsonify({'erro': 'Chave de API não configurada.'}), 400
    caminho_json = OUTPUT_FOLDER / f'{nome}.json'
    if not caminho_json.exists():
        return jsonify({'erro': 'Arquivo não encontrado.'}), 404
    job_id = str(uuid.uuid4())[:8]
    jobs[job_id] = {'status': 'rodando', 'progresso': 0, 'total': 0, 'erro': None}

    def rodar():
        try:
            with open(caminho_json, encoding='utf-8') as f:
                conversas = json.load(f)
            pendentes = [c for c in conversas if not c.get('analisado', False)]
            jobs[job_id]['total'] = len(pendentes)
            def on_progresso(i, resultado):
                jobs[job_id]['progresso'] = i
            analisar_todas(conversas, ANTHROPIC_API_KEY, on_progresso)
            with open(caminho_json, 'w', encoding='utf-8') as f:
                json.dump(conversas, f, ensure_ascii=False, default=str)
            ts = datetime.now().strftime('%Y%m%d_%H%M')
            caminho_excel = OUTPUT_FOLDER / f'{nome}_{ts}.xlsx'
            gerar_excel(conversas, str(caminho_excel))
            jobs[job_id]['status'] = 'concluido'
            jobs[job_id]['excel'] = caminho_excel.name
        except Exception as e:
            jobs[job_id]['status'] = 'erro'
            jobs[job_id]['erro'] = str(e)

    threading.Thread(target=rodar, daemon=True).start()
    return jsonify({'job_id': job_id})


@app.route('/status/<job_id>')
@login_required
def status(job_id):
    job = jobs.get(job_id, {'status': 'nao_encontrado'})
    return jsonify(job)


@app.route('/download/<nome>')
@login_required
def download(nome):
    caminho = OUTPUT_FOLDER / nome
    if not caminho.exists():
        return 'Arquivo não encontrado.', 404
    return send_file(str(caminho), as_attachment=True)


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False)
